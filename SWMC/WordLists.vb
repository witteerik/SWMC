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
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Reflection



Public Enum WordRange
    SubGroup
    WholeLanguage
End Enum

Public Module WordListsIO


    ''' <summary>
    ''' Creates a list of valid phonetic characters using the global arrays SwedishConsonants_IPA, SwedishVowels_IPA, AllSuprasegmentalIPACharacters, ZerpPhoneme, and the global strings AmbiguosOnsetMarker, and AmbiguosCodaMarker
    ''' </summary>
    ''' <returns></returns>
    Public Function CreateListOfValidPhoneticCharactersForSwedish() As List(Of String)

        'Creating a list of valid phonetic characters
        Dim ValidPhoneticCharacters As New List(Of String)
        For p = 0 To SwedishConsonants_IPA.Count - 1
            ValidPhoneticCharacters.Add(SwedishConsonants_IPA(p))
        Next
        For p = 0 To SwedishVowels_IPA.Count - 1
            ValidPhoneticCharacters.Add(SwedishVowels_IPA(p))
        Next
        For p = 0 To AllSuprasegmentalIPACharacters.Count - 1
            ValidPhoneticCharacters.Add(AllSuprasegmentalIPACharacters(p))
        Next

        ValidPhoneticCharacters.Add(ZeroPhoneme)
        ValidPhoneticCharacters.Add(AmbiguosOnsetMarker)
        ValidPhoneticCharacters.Add(AmbiguosCodaMarker)

        SendInfoToLog("Created a list of valid phonetic characters. Result: " & vbCrLf &
            "The following " & ValidPhoneticCharacters.Count & " characters are counted as valid: " & vbCrLf &
                      String.Join(" ", ValidPhoneticCharacters))

        Return ValidPhoneticCharacters

    End Function



    ''' <summary>
    ''' This class holds the order of columns in the input and output word files. The column indices may be changed, but the relative order of the column must be intact.
    ''' </summary>
    Public Class PhoneticTxtStringColumnIndices

        Public Property OrthographicForm As Integer?
        Public Property GIL2P_OT_Average As Integer?
        Public Property GIL2P_OT_Min As Integer?
        Public Property PIP2G_OT_Average As Integer?
        Public Property PIP2G_OT_Min As Integer?
        Public Property G2P_OT_Average As Integer?
        Public Property UpperCase As Integer?
        Public Property Homographs As Integer?
        Public Property HomographCount As Integer?
        Public Property SpecialCharacter As Integer?
        Public Property RawWordTypeFrequency As Integer?
        Public Property RawDocumentCount As Integer?
        Public Property PhoneticForm As Integer?
        Public Property TemporarySyllabification As Integer?
        Public Property ReducedTranscription As Integer?

        Public Property PhonotacticType As Integer?

        Public Property SSPP_Average As Integer?
        Public Property SSPP_Min As Integer?

        Public Property PSP_Sum As Integer?
        Public Property PSBP_Sum As Integer?

        Public Property S_PSP_Average As Integer?
        Public Property S_PSBP_Average As Integer?

        Public Property Homophones As Integer?
        Public Property HomophoneCount As Integer?

        Public Property PNDP As Integer?
        Public Property PLD1Transcriptions As Integer?

        Public Property ONDP As Integer?
        Public Property OLD1Spellings As Integer?

        Public Property PLDx_Average As Integer?
        Public Property OLDx_Average As Integer?
        Public Property PLDx_Neighbors As Integer?
        Public Property OLDx_Neighbors As Integer?

        Public Property Sonographs As Integer?
        Public Property AllPoS As Integer?
        Public Property AllLemmas As Integer?
        Public Property NumberOfSenses As Integer?
        Public Property Abbreviation As Integer?
        Public Property Acronym As Integer?

        Public Property AllPossibleSenses As Integer?
        Public Property SSPP As Integer?
        Public Property PSP As Integer?
        Public Property PSBP As Integer?
        Public Property S_PSP As Integer?
        Public Property S_PSBP As Integer?
        Public Property GIL2P_OT As Integer?
        Public Property PIP2G_OT As Integer?
        Public Property G2P_OT As Integer?
        Public Property ForeignWord As Integer?
        'Public Property CorrectedSpelling As Integer?
        Public Property CorrectedTranscription As Integer?
        'Public Property SAMPA As Integer? 'SAMPA support has been removed
        Public Property ManuallyReveiwedCount As Integer?

        'Variables that are not read from the txt input file, but should be exported to the output file
        Public Property IPA As Integer?
        Public Property ZipfValue As Integer?
        Public Property LetterCount As Integer?
        Public Property GraphemeCount As Integer?
        Public Property DiGraphCount As Integer?
        Public Property TriGraphCount As Integer?
        Public Property LongGraphemesCount As Integer?
        Public Property SyllableCount As Integer?
        Public Property Tone As Integer?
        Public Property MainStressSyllable As Integer?
        Public Property SecondaryStressSyllable As Integer?
        Public Property PhoneCount As Integer?
        Public Property PhoneCountZero As Integer?
        Public Property PLD1WordCount As Integer?
        Public Property OLD1WordCount As Integer?
        Public Property PossiblePoSCount As Integer?
        Public Property MostCommonPoS As Integer?
        Public Property PossibleLemmaCount As Integer?
        Public Property MostCommonLemma As Integer?

        Public Property ManualEvaluations As Integer?
        Public Property ManualEvaluationsCount As Integer?

        Public Property OrthographicIsolationPoint As Integer?
        Public Property PhoneticIsolationPoint As Integer?

        Public Sub New()
            SendInfoToLog("Creating a new instance of :" & Me.ToString)
        End Sub

        ''' <summary>
        ''' Sets a default column order for the import and export of word list files.
        ''' </summary>
        ''' <param name="ColumnsToSkip">May contain column indices that should be excluded from input or output.</param>
        Public Sub SetDefaultOrder(Optional ByRef ColumnsToSkip As List(Of Integer) = Nothing,
                                       Optional ByRef OnlyFinalWordListColumns As Boolean = False)


            SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name &
                              ", OnlyFinalWordListColumns: " & OnlyFinalWordListColumns.ToString)

            Dim TempColumnsToSkip As New List(Of Integer)
            If ColumnsToSkip IsNot Nothing Then
                For n = 0 To ColumnsToSkip.Count - 1
                    TempColumnsToSkip.Add(ColumnsToSkip(n) - n) 'The reason for the -n term is that, if more than one column is excluded, that a value of 1 must be subtracted from total the column index for every excluded column with a lower index, since these previous columns are removed.
                Next
            End If

            Dim ColumnIndex As Integer = 0

            OrthographicForm = ColumnIndex

            GIL2P_OT_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            GIL2P_OT_Min = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PIP2G_OT_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PIP2G_OT_Min = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)

            G2P_OT_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            UpperCase = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            Homographs = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            HomographCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            SpecialCharacter = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            RawWordTypeFrequency = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            RawDocumentCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PhoneticForm = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            'TemporarySyllabification = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            ReducedTranscription = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PhonotacticType = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)

            SSPP_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            SSPP_Min = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)

            PSP_Sum = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PSBP_Sum = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)

            S_PSP_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            S_PSBP_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)

            Homophones = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            HomophoneCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PNDP = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PLD1Transcriptions = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)

            ONDP = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            OLD1Spellings = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)

            PLDx_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            OLDx_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PLDx_Neighbors = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            OLDx_Neighbors = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)

            Sonographs = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            AllPoS = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            AllLemmas = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            NumberOfSenses = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            Abbreviation = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            Acronym = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)

            If OnlyFinalWordListColumns = False Then
                AllPossibleSenses = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
                SSPP = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
                PSP = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
                PSBP = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
                S_PSP = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
                S_PSBP = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            End If

            If OnlyFinalWordListColumns = False Then
                GIL2P_OT = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
                PIP2G_OT = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
                G2P_OT = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            End If

            ForeignWord = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)

            If OnlyFinalWordListColumns = False Then
                'CorrectedSpelling = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
                CorrectedTranscription = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
                'SAMPA = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip) 'SAMPA support has been removed
                ManuallyReveiwedCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            End If

            IPA = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            ZipfValue = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            LetterCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            GraphemeCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            DiGraphCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            TriGraphCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            LongGraphemesCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            SyllableCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            Tone = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            MainStressSyllable = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            SecondaryStressSyllable = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PhoneCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            'PhoneCountZero = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PLD1WordCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            OLD1WordCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PossiblePoSCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            MostCommonPoS = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PossibleLemmaCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            MostCommonLemma = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)

            If OnlyFinalWordListColumns = False Then
                ManualEvaluations = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
                ManualEvaluationsCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            End If

            'Added 2019-08-14
            OrthographicIsolationPoint = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PhoneticIsolationPoint = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)

        End Sub


        Private Function IncreaseColumnIndex(ByRef ColumnIndex As Integer, Optional ByRef TempColumnsToSkip As List(Of Integer) = Nothing) As Integer?

            If Not TempColumnsToSkip Is Nothing Then

                If TempColumnsToSkip.Contains(ColumnIndex + 1) Then
                    TempColumnsToSkip.RemoveAt(0)
                    Return Nothing
                Else
                    ColumnIndex += 1
                    Return ColumnIndex
                End If

            Else
                ColumnIndex += 1
                Return ColumnIndex
            End If

        End Function


        ''' <summary>
        ''' Sets the column order used by the WordMetrics web site.
        ''' </summary>
        Public Sub SetWebSiteColumnOrder(ByRef OnlyWordLevelData As Boolean,
                                         ByRef IncludeOT As Boolean,
                                         ByRef IncludePP As Boolean,
                                         ByRef IncludePND As Boolean,
                                         ByRef IncludeOND As Boolean,
                                         ByRef IncludePLDx As Boolean,
                                         ByRef IncludeOLDx As Boolean,
                                         ByRef IncludeOrthographicIsolationPoints As Boolean,
                                         ByRef IncludePhoneticIsolationPointsCheckBox As Boolean,
                                         Optional ByRef ExtraData As Boolean = False)

            Dim ColumnIndex As Integer = 0

            OrthographicForm = ColumnIndex

            If IncludeOT = True Then
                GIL2P_OT_Average = IncreaseColumnIndex(ColumnIndex)
                GIL2P_OT_Min = IncreaseColumnIndex(ColumnIndex)
                PIP2G_OT_Average = IncreaseColumnIndex(ColumnIndex)
                PIP2G_OT_Min = IncreaseColumnIndex(ColumnIndex)
                G2P_OT_Average = IncreaseColumnIndex(ColumnIndex)
            End If

            If ExtraData = True Then
                UpperCase = IncreaseColumnIndex(ColumnIndex)
                Homographs = IncreaseColumnIndex(ColumnIndex)
            End If
            'HomographCount = IncreaseColumnIndex(ColumnIndex)
            If ExtraData = True Then
                SpecialCharacter = IncreaseColumnIndex(ColumnIndex)
                RawWordTypeFrequency = IncreaseColumnIndex(ColumnIndex)
                RawDocumentCount = IncreaseColumnIndex(ColumnIndex)
            End If

            PhoneticForm = IncreaseColumnIndex(ColumnIndex)
            TemporarySyllabification = IncreaseColumnIndex(ColumnIndex)
            ReducedTranscription = IncreaseColumnIndex(ColumnIndex)

            If ExtraData = True Then PhonotacticType = IncreaseColumnIndex(ColumnIndex)

            If IncludePP = True Then
                SSPP_Average = IncreaseColumnIndex(ColumnIndex)
                SSPP_Min = IncreaseColumnIndex(ColumnIndex)
                PSP_Sum = IncreaseColumnIndex(ColumnIndex)
                PSBP_Sum = IncreaseColumnIndex(ColumnIndex)
                S_PSP_Average = IncreaseColumnIndex(ColumnIndex)
                S_PSBP_Average = IncreaseColumnIndex(ColumnIndex)
            End If

            If ExtraData = True Then Homophones = IncreaseColumnIndex(ColumnIndex)

            'HomophoneCount = IncreaseColumnIndex(ColumnIndex)
            If IncludePND = True Then
                PNDP = IncreaseColumnIndex(ColumnIndex)
                PLD1Transcriptions = IncreaseColumnIndex(ColumnIndex)
            End If

            If IncludeOND = True Then
                ONDP = IncreaseColumnIndex(ColumnIndex)
                OLD1Spellings = IncreaseColumnIndex(ColumnIndex)
            End If

            If IncludePND = True And IncludePLDx = True Then PLDx_Average = IncreaseColumnIndex(ColumnIndex)
            If IncludeOND = True And IncludeOLDx = True Then OLDx_Average = IncreaseColumnIndex(ColumnIndex)

            If IncludePND = True And IncludePLDx = True Then PLDx_Neighbors = IncreaseColumnIndex(ColumnIndex)
            If IncludeOND = True And IncludeOLDx = True Then OLDx_Neighbors = IncreaseColumnIndex(ColumnIndex)

            If IncludeOT = True Then
                Sonographs = IncreaseColumnIndex(ColumnIndex)
            End If

            If ExtraData = True Then
                AllPoS = IncreaseColumnIndex(ColumnIndex)
                AllLemmas = IncreaseColumnIndex(ColumnIndex)
                NumberOfSenses = IncreaseColumnIndex(ColumnIndex)
                Abbreviation = IncreaseColumnIndex(ColumnIndex)
                Acronym = IncreaseColumnIndex(ColumnIndex)
            End If

            'AllPossibleSenses = IncreaseColumnIndex(ColumnIndex)
            If OnlyWordLevelData = False Then
                If IncludePP = True Then
                    SSPP = IncreaseColumnIndex(ColumnIndex)
                    PSP = IncreaseColumnIndex(ColumnIndex)
                    PSBP = IncreaseColumnIndex(ColumnIndex)
                    S_PSP = IncreaseColumnIndex(ColumnIndex)
                    S_PSBP = IncreaseColumnIndex(ColumnIndex)
                End If

                If IncludeOT = True Then
                    GIL2P_OT = IncreaseColumnIndex(ColumnIndex)
                    PIP2G_OT = IncreaseColumnIndex(ColumnIndex)
                    G2P_OT = IncreaseColumnIndex(ColumnIndex)
                End If
            End If

            If ExtraData = True Then ForeignWord = IncreaseColumnIndex(ColumnIndex)

            'CorrectedSpelling = IncreaseColumnIndex(ColumnIndex)
            'CorrectedTranscription = IncreaseColumnIndex(ColumnIndex)
            'SAMPA = IncreaseColumnIndex(ColumnIndex)
            'ManuallyReveiwedCount = IncreaseColumnIndex(ColumnIndex)

            If ExtraData = True Then
                'IPA = IncreaseColumnIndex(ColumnIndex)
                ZipfValue = IncreaseColumnIndex(ColumnIndex)
                LetterCount = IncreaseColumnIndex(ColumnIndex)

                If IncludeOT = True Then
                    GraphemeCount = IncreaseColumnIndex(ColumnIndex)
                    DiGraphCount = IncreaseColumnIndex(ColumnIndex)
                    TriGraphCount = IncreaseColumnIndex(ColumnIndex)
                    LongGraphemesCount = IncreaseColumnIndex(ColumnIndex)
                End If

                SyllableCount = IncreaseColumnIndex(ColumnIndex)
                Tone = IncreaseColumnIndex(ColumnIndex)
                MainStressSyllable = IncreaseColumnIndex(ColumnIndex)
                SecondaryStressSyllable = IncreaseColumnIndex(ColumnIndex)
                PhoneCount = IncreaseColumnIndex(ColumnIndex)
                'PhoneCountZero = IncreaseColumnIndex(ColumnIndex)
            End If

            If IncludePND = True Then PLD1WordCount = IncreaseColumnIndex(ColumnIndex)
            If IncludeOND = True Then OLD1WordCount = IncreaseColumnIndex(ColumnIndex)

            'PossiblePoSCount = IncreaseColumnIndex(ColumnIndex)
            'MostCommonPoS = IncreaseColumnIndex(ColumnIndex)
            'PossibleLemmaCount = IncreaseColumnIndex(ColumnIndex)
            'MostCommonLemma = IncreaseColumnIndex(ColumnIndex)

            ManualEvaluations = IncreaseColumnIndex(ColumnIndex)
            'ManualEvaluationsCount = IncreaseColumnIndex(ColumnIndex)

            'Added 2019-08-14
            If IncludeOrthographicIsolationPoints = True Then
                OrthographicIsolationPoint = IncreaseColumnIndex(ColumnIndex)
            End If
            If IncludePhoneticIsolationPointsCheckBox = True Then
                PhoneticIsolationPoint = IncreaseColumnIndex(ColumnIndex)
            End If

        End Sub


        ''' <summary>
        ''' Sets a column order for the import and export of word list files.
        ''' </summary>
        ''' <param name="ColumnsToSkip">May contain column indices that should be excluded from input or output.</param>
        Public Sub SetAfcListOrder(Optional ByRef ColumnsToSkip As List(Of Integer) = Nothing)


            SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            Dim TempColumnsToSkip As New List(Of Integer)
            If ColumnsToSkip IsNot Nothing Then
                For n = 0 To ColumnsToSkip.Count - 1
                    TempColumnsToSkip.Add(ColumnsToSkip(n) - n) 'The reason for the -n term is that, if more than one column is excluded, that a value of 1 must be subtracted from total the column index for every excluded column with a lower index, since these previous columns are removed.
                Next
            End If

            Dim ColumnIndex As Integer = 0

            OrthographicForm = ColumnIndex
            GIL2P_OT_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            GIL2P_OT_Min = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PIP2G_OT_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PIP2G_OT_Min = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            G2P_OT_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            UpperCase = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            Homographs = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            HomographCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            SpecialCharacter = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            RawWordTypeFrequency = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            RawDocumentCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PhoneticForm = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            TemporarySyllabification = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            ReducedTranscription = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PhonotacticType = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            SSPP_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            SSPP_Min = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PSP_Sum = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PSBP_Sum = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            S_PSP_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            S_PSBP_Average = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            Homophones = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            HomophoneCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PNDP = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PLD1Transcriptions = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            ONDP = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            OLD1Spellings = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            Sonographs = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            AllPoS = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            AllLemmas = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            NumberOfSenses = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            Abbreviation = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            Acronym = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            ForeignWord = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            ZipfValue = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            LetterCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            GraphemeCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            DiGraphCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            TriGraphCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            LongGraphemesCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            SyllableCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            Tone = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            MainStressSyllable = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            SecondaryStressSyllable = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PhoneCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PLD1WordCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            OLD1WordCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PossiblePoSCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)
            PossibleLemmaCount = IncreaseColumnIndex(ColumnIndex, TempColumnsToSkip)

        End Sub



        ''' <summary>
        ''' Returns the number of columns (properties) that should be read
        ''' </summary>
        ''' <returns></returns>
        Public Function GetNumberOfColumns() As Integer

            Dim ColumnOrderProperyInfo() As System.Reflection.PropertyInfo = GetType(PhoneticTxtStringColumnIndices).GetProperties
            'For n = 0 To ColumnOrderProperyInfo.Length - 1
            'If GetType(PhoneticTxtStringColumnIndices).GetProperty(ColumnOrderProperyInfo(n).Name).GetValue(Me) IsNot Nothing Then
            'Dim CurrentValue As Integer = GetType(PhoneticTxtStringColumnIndices).GetProperty(ColumnOrderProperyInfo(n).Name).GetValue(Me)
            'If CurrentValue > Maxvalue Then Maxvalue = CurrentValue
            'End If
            'Next

            Return ColumnOrderProperyInfo.Count

        End Function

        ''' <summary>
        ''' Returns a tab delimited string with the activated column headings.
        ''' </summary>
        ''' <returns></returns>
        Public Function GetColumnHeadingsString()

            Dim HeadingLine As String = ""
            Dim ColumnOrderProperyInfo() As System.Reflection.PropertyInfo = GetType(PhoneticTxtStringColumnIndices).GetProperties
            For n = 0 To ColumnOrderProperyInfo.Length - 1
                If GetType(PhoneticTxtStringColumnIndices).GetProperty(ColumnOrderProperyInfo(n).Name).GetValue(Me) IsNot Nothing Then
                    HeadingLine &= ColumnOrderProperyInfo(n).Name & vbTab
                End If
            Next

            Return HeadingLine

        End Function

        ''' <summary>
        ''' Returns a string containing the string representation of the heading of the specified column index, or "N/A" if the index is outside the bounds of the available columns.
        ''' </summary>
        ''' <returns></returns>
        Public Function GetColumnHeadingsString(ByVal Index As Integer)

            Dim ColumnOrderProperyInfo() As System.Reflection.PropertyInfo = GetType(PhoneticTxtStringColumnIndices).GetProperties

            If GetType(PhoneticTxtStringColumnIndices).GetProperty(ColumnOrderProperyInfo(Index).Name).GetValue(Me) IsNot Nothing Then
                Return ColumnOrderProperyInfo(Index).Name
            Else
                Return "N/A"
            End If

        End Function

    End Class


    '''' <summary>
    '''' Parses the stardard IPA phonetic form into a syllable structure in the referenced word
    '''' </summary>
    '''' <param name="PhoneticInputString">The input IPA form. (Phonetic items should be space delimited. Syllable boundary marker is [.])</param>
    '''' <param name="CurrentWord"></param>
    '''' <param name="ValidPhoneticCharacters"></param>
    '''' <param name="ContainsInvalidPhoneticCharacter"></param>
    '''' <param name="CorrectDoubleSpacesInPhoneticForm"></param>
    '''' <param name="CorrectedDoubleSpacesInPhoneticForm"></param>
    '''' <param name="CheckPhonemeValidity"></param>
    'Public Sub ParseInputPhoneticString(ByRef PhoneticInputString As String, ByRef CurrentWord As Word,
    '                                         ByRef ValidPhoneticCharacters As List(Of String),
    '                                        ByRef ContainsInvalidPhoneticCharacter As Boolean,
    '                                        ByRef CorrectDoubleSpacesInPhoneticForm As Boolean,
    '                                        ByRef CorrectedDoubleSpacesInPhoneticForm As Boolean,
    '                                        ByRef CheckPhonemeValidity As Boolean)


    '    'Only reading phonetic form if the input string is not empty
    '    If Not PhoneticInputString.Trim = "" Then

    '        'Reading the extended syllable array and parsing it to a syllable
    '        Dim ExtendedIpaArraySyllableSplit() As String = {}

    '        ExtendedIpaArraySyllableSplit = PhoneticInputString.Trim(".").Trim.Split(".")

    '        'Replacing any double spaces in the input PhoneticForm with single spaces
    '        If CorrectDoubleSpacesInPhoneticForm = True Then
    '            For s = 0 To ExtendedIpaArraySyllableSplit.Count - 1
    '                If ExtendedIpaArraySyllableSplit(s).Contains("  ") Then
    '                    'Correct the double space (such should not occur)
    '                    ExtendedIpaArraySyllableSplit(s) = ExtendedIpaArraySyllableSplit(s).Replace("  ", " ")
    '                    CorrectedDoubleSpacesInPhoneticForm = True
    '                End If
    '            Next
    '        End If

    '        'Checking that all phonetic characters are valid
    '        If CheckPhonemeValidity = True And ValidPhoneticCharacters IsNot Nothing Then
    '            For s = 0 To ExtendedIpaArraySyllableSplit.Length - 1
    '                Dim SyllSplit() As String = ExtendedIpaArraySyllableSplit(s).Trim.Split(" ")
    '                For p = 0 To SyllSplit.Length - 1
    '                    If Not ValidPhoneticCharacters.Contains(SyllSplit(p)) Then
    '                        ContainsInvalidPhoneticCharacter = True
    '                    End If
    '                Next
    '            Next
    '        End If


    '        'Reading suprasegmentals (Tone and index of primary and secondary stress) from the ExtendedIpaArray
    '        'Setting default values
    '        CurrentWord.Syllables.Tone = 0
    '        CurrentWord.Syllables.MainStressSyllableIndex = 0
    '        CurrentWord.Syllables.SecondaryStressSyllableIndex = 0

    '        'Reading values
    '        For syllable = 0 To ExtendedIpaArraySyllableSplit.Count - 1

    '            'Detecting primary stress and tone 1
    '            If ExtendedIpaArraySyllableSplit(syllable).Contains(SpecialCharacters.IpaMainStress) Then
    '                CurrentWord.Syllables.Tone = 1
    '                CurrentWord.Syllables.MainStressSyllableIndex = syllable + 1
    '            End If

    '            'Detecting primary stress and tone 2
    '            If ExtendedIpaArraySyllableSplit(syllable).Contains(SpecialCharacters.IpaMainSwedishAccent2) Then
    '                CurrentWord.Syllables.Tone = 2
    '                CurrentWord.Syllables.MainStressSyllableIndex = syllable + 1
    '            End If

    '            'Detecting secondary stress
    '            If ExtendedIpaArraySyllableSplit(syllable).Contains(SpecialCharacters.IpaSecondaryStress) Then
    '                CurrentWord.Syllables.SecondaryStressSyllableIndex = syllable + 1
    '            End If
    '        Next

    '        'Also copying the suprasegmentals to the word variebles
    '        CurrentWord.Tone = CurrentWord.Syllables.Tone
    '        CurrentWord.MainStressSyllableIndex = CurrentWord.Syllables.MainStressSyllableIndex
    '        CurrentWord.SecondaryStressSyllableIndex = CurrentWord.Syllables.SecondaryStressSyllableIndex


    '        For syllable = 0 To ExtendedIpaArraySyllableSplit.Count - 1

    '            Dim newSyllable As New Word.Syllable

    '            'Sets syllable suprasegmentals
    '            If (syllable = CurrentWord.Syllables.MainStressSyllableIndex - 1 And CurrentWord.Syllables.MainStressSyllableIndex <> 0) Then
    '                newSyllable.IsStressed = True
    '            End If
    '            If (syllable = CurrentWord.Syllables.SecondaryStressSyllableIndex - 1 And CurrentWord.Syllables.SecondaryStressSyllableIndex <> 0) Then
    '                newSyllable.IsStressed = True
    '                newSyllable.CarriesSecondaryStress = True
    '            End If

    '            Dim SyllableArraySplit() As String = ExtendedIpaArraySyllableSplit(syllable).Trim(" ").Split(" ")
    '            For phoneme = 0 To SyllableArraySplit.Count - 1
    '                'Adding only phoneme characters
    '                If Not AllSuprasegmentalIPACharacters.Contains(SyllableArraySplit(phoneme)) Then
    '                    newSyllable.Phonemes.Add(SyllableArraySplit(phoneme).Trim)
    '                End If
    '            Next

    '            'Detecting and removing ambigous syllable markers
    '            If newSyllable.Phonemes(0).Contains(AmbiguosOnsetMarker) Then
    '                newSyllable.AmbigousOnset = True
    '                newSyllable.Phonemes(0) = newSyllable.Phonemes(0).Replace(AmbiguosOnsetMarker, "")
    '            End If
    '            If newSyllable.Phonemes(0).Contains(AmbiguosCodaMarker) Then
    '                newSyllable.AmbigousCoda = True
    '                newSyllable.Phonemes(0) = newSyllable.Phonemes(0).Replace(AmbiguosCodaMarker, "")
    '            End If

    '            CurrentWord.Syllables.Add(newSyllable)
    '        Next
    '    End If


    'End Sub


End Module

''' <summary>
''' A class for storage, manipulation and analyses of groups of Word.
''' </summary>
<Serializable>
Public Class WordGroup

#Region "Declarations"



    Public Property GroupHomophoneCount As Integer
    Public Property GroupHomographCount As Integer
    Public Property LanguageHomophoneCount As Integer
    Public Property LanguageHomographCount As Integer

    'Corpus description
    Public Property CorpusTokenCount As Long
    Public Property CorpusWordTypeCount As Long
    Public Property CorpusDocumentCount As Integer
    Public Property CorpusSentenceCount As Long

    'For minimal-variation-of-phonemes groups
    Public Property MiminalVariationSegmentIndex As SByte
    Public Property TestSyllableIndex As SByte
    Public Property IndexOfTestSyllableNucleus As SByte
    Public Property LengthOfTestSyllableOnset As SByte
    Public Property LengthOfTestSyllableCoda As SByte
    Public Property PreMVSSegments As String
    Public Property PostMVSSegments As String
    Public Property ContrastingNeutralizedPhonemes As Integer

    'Lists
    Public Property PhonemeInventory As List(Of String)
    Public Property MiminalVariationSegments As New List(Of String)
    Public Property GraphemeInventory As List(Of String)

    Private _MemberWords As New List(Of Word)
    Public Property MemberWords As List(Of Word)
        Get
            If _MemberWords Is Nothing Then _MemberWords = New List(Of Word)
            Return _MemberWords
        End Get
        Set(value As List(Of Word))
            'MemberCount = value.Count
            _MemberWords = value
        End Set
    End Property

    Public Property PhonemeDistribution As DataTable

    'Variable for holding a value indicateing the progress upon saving to file, so that the progress can be resumed at that point
    Public Property CurrentWordIndex As Integer = 0

#End Region



#Region "Phonology"



    Public Sub RemoveDoublePhoneticCharacters(Optional ByVal RemoveConsonants As Boolean = True, Optional ByVal RemoveVowels As Boolean = False,
                                                  Optional LogWords As Boolean = True, Optional ExcludeForeignWords As Boolean = True,
                                                  Optional ExcludeAbbreviations As Boolean = True, Optional ByVal DisregardPhoneticLengthCharacter As Boolean = False,
                                               Optional IncludeSAMPA As Boolean = True)

        'Removing all double consonants found (which would be erraneous)
        SendInfoToLog("     Starting to remove double phonetic characters." & vbCrLf &
                          "     Settings:  " & vbCrLf &
                          "     RemoveConsonants:  " & RemoveConsonants.ToString & vbCrLf &
                          "     RemoveVowels:  " & RemoveVowels.ToString & vbCrLf &
                          "     ExcludeForeignWords:  " & ExcludeForeignWords.ToString & vbCrLf &
                          "     ExcludeAbbreviations:  " & ExcludeAbbreviations.ToString & vbCrLf &
                          "     LogWords:  " & LogWords.ToString)

        Dim WordsWithRemovedPhonemes As Integer = 0
        Dim TotalRemovedPhonemesCount As Integer = 0

        Dim ExcludedForeignWords As Integer = 0
        Dim ExcludedAbbreviations As Integer = 0

        For word = 0 To MemberWords.Count - 1

            If ExcludeForeignWords = True Then
                If MemberWords(word).ForeignWord = True Then
                    ExcludedForeignWords += 1
                    Continue For
                End If
            End If

            If ExcludeAbbreviations = True Then
                If MemberWords(word).Abbreviation = True Then
                    ExcludedAbbreviations += 1
                    Continue For
                End If
            End If


            Dim CurrentWordRemovedPhonemesCount As Integer = MemberWords(word).RemoveDoublePhoneticCharacters(RemoveConsonants, RemoveVowels, LogWords, DisregardPhoneticLengthCharacter, IncludeSAMPA)
            If CurrentWordRemovedPhonemesCount > 0 Then
                WordsWithRemovedPhonemes += 1
                TotalRemovedPhonemesCount += CurrentWordRemovedPhonemesCount
            End If

        Next

        SendInfoToLog("     Removed " & TotalRemovedPhonemesCount & " Double consonant phonemes In " & WordsWithRemovedPhonemes & " words." & vbCrLf &
                              "        Excluded " & ExcludedForeignWords & " foreign words from analysis." & vbCrLf &
                              "        Excluded " & ExcludedAbbreviations & " abbreviations from analysis.")


    End Sub


    Public Enum LongConsonantPositions
        SyllableCoda
        SyllableOnset
        Reduplicated
    End Enum
    Public Sub ChangeAmbiSyllabicLongConsonantPositions(ByRef SetLongConsonantPositions As LongConsonantPositions, Optional ByVal StoreOldSyllabificationInAlternative As Boolean = True)

        SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " Setting: " & SetLongConsonantPositions.ToString)

        'Starting a progress window
        Dim myProgress As New ProgressDisplay
        myProgress.Initialize(MemberWords.Count - 1, 0, "Changing ambisyllabic long consonant positions...", 100)
        myProgress.Show()

        For word = 0 To MemberWords.Count - 1

            'Updating progress
            myProgress.UpdateProgress(word)

            If StoreOldSyllabificationInAlternative = True Then
                MemberWords(word).Syllables_AlternateSyllabification = MemberWords(word).Syllables.CreateCopy
            End If

            For syll = 0 To MemberWords(word).Syllables.Count - 1

                Select Case SetLongConsonantPositions
                    Case LongConsonantPositions.SyllableOnset

                        If Not syll = MemberWords(word).Syllables.Count - 1 Then

                            'Moving any long consonats from the last coda position to the first onset position in the next syllable
                            'Only if either a long or short version of the same consonant do not already exists in that position
                            Dim CurrentSyllableLength As Integer = MemberWords(word).Syllables(syll).Phonemes.Count
                            Dim LastPhoneme As String = MemberWords(word).Syllables(syll).Phonemes(CurrentSyllableLength - 1)
                            If SwedishConsonants_IPA.Contains(LastPhoneme) And LastPhoneme.Contains(PhoneticLength) Then
                                If MemberWords(word).Syllables(syll + 1).Phonemes(0) <> LastPhoneme.Replace(PhoneticLength, "") Or
                                    MemberWords(word).Syllables(syll + 1).Phonemes(0) <> LastPhoneme.Replace(PhoneticLength, "") Then

                                    MemberWords(word).Syllables(syll + 1).Phonemes.Insert(0, LastPhoneme)
                                    MemberWords(word).Syllables(syll).Phonemes.RemoveAt(CurrentSyllableLength - 1)
                                    MemberWords(word).DetermineSyllableIndices()

                                End If
                            End If
                        End If


                    Case LongConsonantPositions.SyllableCoda, LongConsonantPositions.Reduplicated

                        If Not syll = 0 Then

                            'Moving any long consonats from the first onset position to the last coda position in the previous syllable
                            'Only if either a long or short version of the same consonant do not already exists in that position
                            Dim FirstPhoneme As String = MemberWords(word).Syllables(syll).Phonemes(0)
                            If SwedishConsonants_IPA.Contains(FirstPhoneme) And FirstPhoneme.Contains(PhoneticLength) Then
                                If MemberWords(word).Syllables.GetPreviousSound(syll, 0) <> FirstPhoneme.Replace(PhoneticLength, "") Or
                                    MemberWords(word).Syllables.GetPreviousSound(syll, 0) <> FirstPhoneme.Replace(PhoneticLength, "") Then

                                    MemberWords(word).Syllables(syll - 1).Phonemes.Add(FirstPhoneme)
                                    MemberWords(word).Syllables(syll).Phonemes.RemoveAt(0)
                                    MemberWords(word).DetermineSyllableIndices()

                                End If
                            End If
                        End If
                End Select
            Next
        Next

        If SetLongConsonantPositions = LongConsonantPositions.Reduplicated Then

            For word = 0 To MemberWords.Count - 2

                'Updating progress
                myProgress.UpdateProgress(word)

                For syll = 0 To MemberWords(word).Syllables.Count - 1

                    'First all are moved to coda above, then reduplication is done from there
                    'Copying a short version of the long last coda consonant to the next syllable

                    If Not syll = MemberWords(word).Syllables.Count - 1 Then

                        'Only if either a long or short version of the same consonant do not already exists in that position
                        'Also only if there is a vowel on the first position in the next syllable (otherwise the cluster is not ambisyllabic)
                        Dim CurrentSyllableLength As Integer = MemberWords(word).Syllables(syll).Phonemes.Count
                        Dim LastPhoneme As String = MemberWords(word).Syllables(syll).Phonemes(CurrentSyllableLength - 1)
                        If SwedishConsonants_IPA.Contains(LastPhoneme) And LastPhoneme.Contains(PhoneticLength) Then

                            If SwedishVowels_IPA.Contains(MemberWords(word).Syllables(syll + 1).Phonemes(0)) Then
                                MemberWords(word).Syllables(syll + 1).Phonemes.Insert(0, LastPhoneme.Replace(PhoneticLength, ""))
                                MemberWords(word).DetermineSyllableIndices()
                            End If

                        End If
                    End If
                Next
            Next
        End If

        'Closing the progress display
        myProgress.Close()

        SendInfoToLog("Method " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " completed successfully.")

    End Sub


    ''' <summary>
    ''' Determining internal syllable structure of all syllables of all words in the word group
    ''' </summary>
    Public Sub DetermineSyllableIndices(Optional ByVal ExcludeAbbreviations As Boolean = False, Optional ByVal ExcludeForeignWords As Boolean = False)

        Dim TotalErrorCount As Integer = 0
        Dim ExcludedAbbreviations As Integer = 0
        Dim ExcludedForeignWords As Integer = 0

        SendInfoToLog("Determining internal syllable structure of all words, and looks for errors in syllable structure." & vbCrLf &
                          "         Settings: " & vbCrLf &
                          "         ExcludeAbbreviations: " & ExcludeAbbreviations & vbCrLf &
                          "         ExcludeForeignWords: " & ExcludeForeignWords)

        For word = 0 To MemberWords.Count - 1
            If ExcludeAbbreviations = True Then
                If MemberWords(word).Abbreviation = True Then
                    ExcludedAbbreviations += 1
                    Continue For
                End If
            End If

            If ExcludeForeignWords = True Then
                If MemberWords(word).ForeignWord = True Then
                    ExcludedForeignWords += 1
                    Continue For
                End If
            End If

            If MemberWords(word).DetermineSyllableIndices() > 0 Then TotalErrorCount += 1

        Next

        SendInfoToLog("     Finished analysing internal syllable structure. " & TotalErrorCount & " words with errors in syllable structure were found." & vbCrLf &
                          "        Excluded " & ExcludedAbbreviations & " abbreviation words from analysis." & vbCrLf &
                          "        Excluded " & ExcludedForeignWords & " foreign words from analysis.")

    End Sub


    ''' <summary>
    ''' Looks for errors in syllable structure of all words in the word list.
    ''' </summary>
    Public Sub MarkSyllableWeightErrors(Optional ByRef ExcludeAbbreviations As Boolean = True,
                                                Optional ByRef ExcludeForeignWords As Boolean = True,
                                             Optional LogErrorWords As Boolean = False,
                                             Optional LogDetailedErrorInfo As Boolean = False)

        SendInfoToLog("Looks for errors in syllable Weight of all words in the word list.")

        Dim TotalErrorCount As Integer = 0
        Dim ExcludedAbbreviations As Integer = 0
        Dim ExcludedForeignWords As Integer = 0

        'Starting a progress window
        Dim myProgress As New ProgressDisplay
        myProgress.Initialize(MemberWords.Count - 1, 0, "Looks for syllable Weight errors...", 100)
        myProgress.Show()

        For word = 0 To MemberWords.Count - 1

            'Updating progress
            myProgress.UpdateProgress(word)

            If ExcludeAbbreviations = True And MemberWords(word).Abbreviation = True Then
                ExcludedAbbreviations += 1
                Continue For
            End If

            If ExcludeForeignWords = True And MemberWords(word).ForeignWord = True Then
                ExcludedForeignWords += 1
                Continue For
            End If

            If MemberWords(word).MarkSyllableWeightErrors(LogDetailedErrorInfo) > 0 Then
                TotalErrorCount += 1
                If LogErrorWords = True Then SendInfoToLog(String.Join(" ", MemberWords(word).BuildExtendedIpaArray), "SyllableWeightErrorWords")
            End If
        Next

        'Closing the progress display
        myProgress.Close()

        SendInfoToLog("     Finished looking for errors in syllable Weight. " & TotalErrorCount & " words with erros found." & vbCrLf &
                          "         Excluded " & ExcludedAbbreviations & " abbreviation words from analysis." & vbCrLf &
                          "         Excluded " & ExcludedForeignWords & " foreign words from analysis.")


    End Sub

    Public Sub MarkPhoneticLengthInWrongPlace(Optional ByRef ExcludeAbbreviations As Boolean = True,
                                                Optional ByRef ExcludeForeignWords As Boolean = True)

        SendInfoToLog("Looks for length markings in wrong places.")

        Dim TotalErrorCount As Integer = 0
        Dim ExcludedAbbreviations As Integer = 0
        Dim ExcludedForeignWords As Integer = 0

        'Starting a progress window
        Dim myProgress As New ProgressDisplay
        myProgress.Initialize(MemberWords.Count - 1, 0, "Looks for length errors...", 100)
        myProgress.Show()

        For word = 0 To MemberWords.Count - 1

            'Updating progress
            myProgress.UpdateProgress(word)

            If ExcludeAbbreviations = True And MemberWords(word).Abbreviation = True Then
                ExcludedAbbreviations += 1
                Continue For
            End If

            If ExcludeForeignWords = True And MemberWords(word).ForeignWord = True Then
                ExcludedForeignWords += 1
                Continue For
            End If

            If MemberWords(word).MarkPhoneticLengthInWrongPlace() > 0 Then
                TotalErrorCount += 1
            End If
        Next

        'Closing the progress display
        myProgress.Close()

        SendInfoToLog("     Finished looking for length errors. " & TotalErrorCount & " words with length errors found." & vbCrLf &
                          "         Excluded " & ExcludedAbbreviations & " abbreviation words from analysis." & vbCrLf &
                          "         Excluded " & ExcludedForeignWords & " foreign words from analysis.")

    End Sub


    ''' <summary>
    ''' Determines syllable openness of all syllables of all words in the word list.
    ''' </summary>
    Public Sub DetermineSyllableOpenness(Optional ByVal ExcludeAbbreviations As Boolean = False, Optional ByVal ExcludeForeignWords As Boolean = False)

        SendInfoToLog("Analyses syllable openness in all syllables of all words in the word list." & vbCrLf &
                          "         Settings: " & vbCrLf &
                          "         ExcludeAbbreviations: " & ExcludeAbbreviations & vbCrLf &
                          "         ExcludeForeignWords: " & ExcludeForeignWords)

        Dim WordsWithAmbigousSyllableBoundaries As Integer = 0
        Dim ExcludedAbbreviations As Integer = 0
        Dim ExcludedForeignWords As Integer = 0

        For word = 0 To MemberWords.Count - 1

            If ExcludeAbbreviations = True Then
                If MemberWords(word).Abbreviation = True Then
                    ExcludedAbbreviations += 1
                    Continue For
                End If
            End If

            If ExcludeForeignWords = True Then
                If MemberWords(word).ForeignWord = True Then
                    ExcludedForeignWords += 1
                    Continue For
                End If
            End If

            If MemberWords(word).DetectAmbigousSyllableBoundaries() = True Then WordsWithAmbigousSyllableBoundaries += 1
            MemberWords(word).DetermineSyllableOpenness()
        Next

        SendInfoToLog("          Finished analysis of syllable openness. Detected " & WordsWithAmbigousSyllableBoundaries & " words with ambigous syllable boundaries." & vbCrLf &
                          "            Excluded " & ExcludedAbbreviations & " abbreviation words from analysis." & vbCrLf &
                          "            Excluded " & ExcludedForeignWords & " foreign words from analysis.")


    End Sub



    ''' <summary>
    ''' Returns the length of the longest ExtendedIPA form stored in any member word, or -1 if wordlist is empty.
    ''' </summary>
    ''' <returns></returns>
    Public Function GetLongestTranscriptionStringLength() As Integer

        Dim MaxLength As Integer = -1

        'Getting the longest ExtendedIPA form
        For word = 0 To MemberWords.Count - 1
            If MemberWords(word).TranscriptionString.Length > MaxLength Then MaxLength = MemberWords(word).TranscriptionString.Length
        Next

        Return MaxLength

    End Function



    ''' <summary>
    ''' Returns the highest syllable count of any member word, or -1 if wordlist is empty.
    ''' </summary>
    ''' <returns></returns>
    Public Function GetHighestSyllableCount() As Integer

        Dim MaxLength As Integer = -1

        'Getting the longest ExtendedIPA form
        For word = 0 To MemberWords.Count - 1
            If MemberWords(word).Syllables.Count > MaxLength Then MaxLength = MemberWords(word).Syllables.Count
        Next

        Return MaxLength

    End Function


#End Region

#Region "Orthography"


    ''' <summary>
    ''' Returns the length of the longest orthographic form stored in any member word, or -1 if wordlist is empty.
    ''' </summary>
    ''' <returns></returns>
    Public Function GetLongestOrthographicFormLength() As Integer

        Dim MaxLength As Integer = -1

        'Getting the longest orthographic form
        For word = 0 To MemberWords.Count - 1
            If MemberWords(word).OrthographicForm.Length > MaxLength Then MaxLength = MemberWords(word).OrthographicForm.Length
        Next

        Return MaxLength

    End Function


#End Region


#Region "FrequencyDistributions"


    ''' <summary>
    ''' For each word length (phoneme count) occuring in the word list, calculates the summed raw frequency of occurence based on the raw word type frequency data.
    ''' </summary>
    Public Sub WordLengthCommonality()

        SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

        Dim SumList As New SortedList(Of Integer, Integer) 'Key: word length, Value: frequency of occurence

        'Counting
        For Each MemberWord In MemberWords

            Dim PhonemeCount As Integer = MemberWord.CountPhonemes
            If MemberWord.RawWordTypeFrequency > 0 Then
                If Not SumList.ContainsKey(PhonemeCount) Then
                    SumList.Add(PhonemeCount, MemberWord.RawWordTypeFrequency)
                Else
                    SumList(PhonemeCount) += MemberWord.RawWordTypeFrequency
                End If
            End If

        Next

        'Exporting data
        Dim OutputString As String = "Frequency of various Word lengths (phoneme count):" & vbCrLf & "WordLength" & vbTab & "Occurences" & vbCrLf

        For Each CurrentItem In SumList
            OutputString &= CurrentItem.Key & vbTab & CurrentItem.Value & vbCrLf
        Next

        SendInfoToLog(vbCrLf & OutputString, "WordLengthOccurences")

        SendInfoToLog("Method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " completed successfully.")

    End Sub



    ''' <summary>
    ''' Calculated Zipf-values of all words in the WordGroup
    ''' </summary>
    Public Sub CalculateZipfValues(Optional ByRef PositionTerm As Integer = 3)

        SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

        SendInfoToLog("    Values used for Zipf value calculation: CorpusTokenCount: " & CorpusTokenCount & "; CorpusWordTypeCount: " & CorpusWordTypeCount)

        'Tests if the word group corpus description data is assigned (not 0)
        If CorpusTokenCount = 0 Or CorpusWordTypeCount = 0 Then
            MsgBox("CorpusTokenCount or CorpusWordTypeCount has a value of 0. Incorrect ZipfValue will be calculated.")
            SendInfoToLog("Warning: Zipfvalue is calculated using incorrect values for CorpusTokenCount (" & CorpusTokenCount & ") or CorpusWordTypeCount (" & CorpusWordTypeCount & vbCrLf &
                                  "None of these should be 0.")
        End If

        For word = 0 To MemberWords.Count - 1

            MemberWords(word).ZipfValue_Word = CalculateZipfValue(MemberWords(word).RawWordTypeFrequency,
                                                                          CorpusTokenCount, CorpusWordTypeCount, PositionTerm)

        Next

        SendInfoToLog("Method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " completed successfully. Results: Zipf values was calculated for all " & MemberWords.Count & " words.")

    End Sub


#End Region

#Region "CorpusMethods"



    ''' <summary>
    ''' Copies the corpus description data (CorpusTokenCount, CorpusWordTypeCount, CorpusDocumentCount and CorpusSentenceCount) from another word group.
    ''' </summary>
    ''' <param name="SourceWordgroup"></param>
    Public Sub GetCorpusInfoFromOtherWordgroup(ByRef SourceWordgroup As WordGroup)

        Me.CorpusTokenCount = SourceWordgroup.CorpusTokenCount
        Me.CorpusWordTypeCount = SourceWordgroup.CorpusWordTypeCount
        Me.CorpusDocumentCount = SourceWordgroup.CorpusDocumentCount
        Me.CorpusSentenceCount = SourceWordgroup.CorpusSentenceCount

    End Sub


#End Region

#Region "Sonographs"



    Public Class CountExamples
        Property Count As Integer
        Property Examples As New List(Of String)
    End Class

    Public Enum InclusionChoice
        Exclude
        Include
    End Enum

    ''' <summary>
    ''' Matches the phonetic transcription to the orthographic form using a set of rules specified in a .txt file.
    ''' </summary>
    ''' <param name="InclusiveWordClassDetectionCodes">May contain a list of word class codes to be identified. Any word with at least one of the word class assignment in the list, will either be included or excluded depending on the parameter InclusiveWordClassDetectionRole.</param>
    ''' <param name="InclusiveWordClassDetectionRole">Determines what should happen (exclusion or inclusion) to the words identified by InclusiveWordClassDetectionCodes.</param>
    ''' <param name="ExclusiveWordClassDetectionCodes">May contain a list of word class codes to be identified. Any word containing ONLY ONE (and no other) of the word class assignments in the list will be identified, and then either included or excluded depending on the parameter InclusiveWordClassDetectionRole.</param>
    ''' <param name="ExclusiveWordClassDetectionRole">Determines what should happen (exclusion or inclusion) to the words identified by ExclusiveWordClassDetectionCodes.</param>
    ''' <param name="ExclusiveWordClassCombinationDetectionCodes">May contain a list of word class codes to be identified. Any word containing any of the word class assignments, but no other word class assignments, in the list, will either be included or excluded depending on the parameter InclusiveWordClassDetectionRole.</param>
    ''' <param name="ExclusiveWordClassCombinationDetectionRole">Determines what should happen (exclusion or inclusion) to the words identified by ExclusiveWordClassCombinationDetectionCodes.</param>
    ''' <returns>Returns the PhoneToGraphemesDictionary used.</returns>
    Public Function GetSonographs(ByRef p2g_Settings As p2gParameters,
                                                         Optional ByVal InclusiveWordClassDetectionCodes() As String = Nothing,
                                                         Optional ByVal InclusiveWordClassDetectionRole As InclusionChoice = InclusionChoice.Exclude,
                                                         Optional ByVal ExclusiveWordClassDetectionCodes() As String = Nothing,
                                                         Optional ByVal ExclusiveWordClassDetectionRole As InclusionChoice = InclusionChoice.Exclude,
                                                         Optional ByVal ExclusiveWordClassCombinationDetectionCodes() As String = Nothing,
                                                         Optional ByVal ExclusiveWordClassCombinationDetectionRole As InclusionChoice = InclusionChoice.Exclude,
                                                         Optional ByVal ExcludeAbbreviations As Boolean = False,
                                                         Optional ByVal ExcludeForeignWords As Boolean = False) As PhoneToGraphemes

        Dim InclusiveWordClassDetectionExclusionRole As Boolean = True
        If InclusiveWordClassDetectionRole = InclusionChoice.Include Then InclusiveWordClassDetectionExclusionRole = False
        If InclusiveWordClassDetectionCodes Is Nothing Then InclusiveWordClassDetectionCodes = {}

        Dim ExclusiveWordClassDetectionExclusionRole As Boolean = True
        If ExclusiveWordClassDetectionRole = InclusionChoice.Include Then ExclusiveWordClassDetectionExclusionRole = False
        If ExclusiveWordClassDetectionCodes Is Nothing Then ExclusiveWordClassDetectionCodes = {}

        Dim ExclusiveWordClassCombinationDetectionExclusionRole As Boolean = True
        If ExclusiveWordClassCombinationDetectionRole = InclusionChoice.Include Then ExclusiveWordClassCombinationDetectionExclusionRole = False
        If ExclusiveWordClassCombinationDetectionCodes Is Nothing Then ExclusiveWordClassCombinationDetectionCodes = {}

        'Setting the default output folder
        If p2g_Settings.OutputFolder = "" Then p2g_Settings.OutputFolder = logFilePath

        Dim p2g_SettingsLogString As String = p2g_Settings.CreateSettingsLogString

        SendInfoToLog("Initializing GetSonographs of all " & MemberWords.Count & " words in the input word list." & vbCrLf &
                          "         Settings:  " & vbCrLf &
                          "         InclusiveWordClassDetectionCodes  " & String.Join(" ", InclusiveWordClassDetectionCodes) & vbCrLf &
                          "         InclusiveWordClassDetectionRole  " & InclusiveWordClassDetectionRole.ToString & vbCrLf &
                          "         ExclusiveWordClassDetectionCodes  " & String.Join(" ", ExclusiveWordClassDetectionCodes) & vbCrLf &
                          "         ExclusiveWordClassDetectionRole  " & ExclusiveWordClassDetectionRole.ToString & vbCrLf &
                          "         ExclusiveWordClassCombinationDetectionCodes  " & String.Join(" ", ExclusiveWordClassCombinationDetectionCodes) & vbCrLf &
                          "         ExclusiveWordClassCombinationDetectionRole  " & ExclusiveWordClassCombinationDetectionRole.ToString & vbCrLf &
                          "         ExcludeForeignWords  " & ExcludeForeignWords.ToString & vbCrLf &
                          "         Current p2g_settings  : " & vbCrLf & p2g_SettingsLogString)


        p2g_Settings.ErrorDictionary = New Dictionary(Of String, CountExamples)
        p2g_Settings.ListOfGraphemes = New List(Of String)

        Dim ExcludedAbbreviations As Integer = 0
        Dim ExcludedForeignWords As Integer = 0

        Dim Excluded_InclusiveWordClassDetectionWords As Integer = 0
        Dim Included_InclusiveWordClassDetectionWords As Integer = 0

        Dim Excluded_ExclusiveWordClassDetectionWords As Integer = 0
        Dim Included_ExclusiveWordClassDetectionWords As Integer = 0

        Dim Excluded_ExclusiveWordClassCombinationDetectionWords As Integer = 0
        Dim Included_ExclusiveWordClassCombinationDetectionWords As Integer = 0

        Dim ProcessedWordsCount As Integer = 0


        'Calculates the sampling points for example words
        Dim ExampleWordSamplingPoints As New SortedSet(Of Integer)
        If p2g_Settings.UseExampleWordSampling Then
            Dim SamplingInterval As Integer = (MemberWords.Count - 1 - (MemberWords.Count / 10)) / p2g_Settings.MaximumNumberOfExampleWords
            For n = 0 To p2g_Settings.MaximumNumberOfExampleWords - 1
                ExampleWordSamplingPoints.Add(n * SamplingInterval)
            Next
        End If

        'Starting a progress window
        Dim myProgress As New ProgressDisplay
        myProgress.Title = "p2g in progress..."
        myProgress.Initialize(MemberWords.Count - 1, 0, "p2g in progress...", 100)
        myProgress.Show()

        For word = 0 To MemberWords.Count - 1

            'Updating progress
            myProgress.UpdateProgress(word)

            'Activating example sampling on the sampling points
            If p2g_Settings.UseExampleWordSampling Then
                If ExampleWordSamplingPoints.Contains(word) Then
                    p2g_Settings.PhoneToGraphemesDictionary.ActivateExampleSampling()
                End If
            End If

            'Word exclusions / inclusions
            If ExcludeAbbreviations = True Then
                If MemberWords(word).Abbreviation = True Then
                    ExcludedAbbreviations += 1
                    Continue For
                End If
            End If

            If ExcludeForeignWords = True Then
                If MemberWords(word).ForeignWord = True Then
                    ExcludedForeignWords += 1
                    Continue For
                End If
            End If

            'Inclusive word class detection
            If InclusiveWordClassDetectionCodes.Count > 0 Then
                If MemberWords(word).DetectWordClass(InclusiveWordClassDetectionCodes) = InclusiveWordClassDetectionExclusionRole Then
                    Excluded_InclusiveWordClassDetectionWords += 1
                    Continue For
                Else
                    Included_InclusiveWordClassDetectionWords += 1
                End If
            End If

            'Exclusive word class detection
            If ExclusiveWordClassDetectionCodes.Length > 0 Then
                Dim Abort1 As Boolean = False
                For n = 0 To ExclusiveWordClassDetectionCodes.Count - 1
                    If MemberWords(word).WordIsOnlyOneWordClass(ExclusiveWordClassDetectionCodes(n)) = ExclusiveWordClassDetectionExclusionRole Then
                        Abort1 = True
                        Exit For
                    End If
                Next
                If Abort1 = True Then
                    Excluded_ExclusiveWordClassDetectionWords += 1
                    Continue For
                Else
                    Included_ExclusiveWordClassDetectionWords += 1
                End If
            End If

            'Exclusive word class combination detection (TODO: NB: This function has not yet been properly tested (2016-09-20)!)
            If ExclusiveWordClassCombinationDetectionCodes.Count > 0 Then
                If MemberWords(word).ExcludeWordDueToWordClasses(ExclusiveWordClassCombinationDetectionCodes) = ExclusiveWordClassCombinationDetectionExclusionRole Then
                    Excluded_ExclusiveWordClassCombinationDetectionWords += 1
                    Continue For
                Else
                    Included_ExclusiveWordClassCombinationDetectionWords += 1
                End If
            End If


            'Counting inculded words
            ProcessedWordsCount += 1

            'Running p2g parsing of the current word
            MemberWords(word).Generate_p2g_Data(p2g_Settings)

        Next word


        'Updating progress
        myProgress.Title = "p2g Completed. Saving data."

        'Exporting graphemes
        SendInfoToLog(vbCrLf & String.Join(vbCrLf, p2g_Settings.ListOfGraphemes), "Graphemes", p2g_Settings.OutputFolder)

        If p2g_Settings.DontExportAnything = False Then
            'Exporting the p2g dictionary used (in a reusable format)
            p2g_Settings.PhoneToGraphemesDictionary.Export_p2g_Examples(logFilePath, "p2g_LastUsedVersion")

            'Exporting the p2g dictionary used (in a reusable format), containing only the most common example words
            p2g_Settings.PhoneToGraphemesDictionary.Export_p2g_MostCommonExample(logFilePath, "p2g_LastUsedVersion_SelectedExampleWords")
        End If

        'Exporting the error dictionary
        Dim ErrorDictionaryOutput As String = vbCrLf
        For Each item In p2g_Settings.ErrorDictionary
            ErrorDictionaryOutput &= item.Key & vbTab & item.Value.Count & vbTab & String.Join(" | ", item.Value.Examples).Replace(p2g_Settings.PhoneToGraphemesDictionary.WordStartMarker, "").Replace(p2g_Settings.PhoneToGraphemesDictionary.WordEndMarker, "") & vbCrLf
        Next
        SendInfoToLog(ErrorDictionaryOutput, "ErrorDictionaryOutput", p2g_Settings.OutputFolder)



        'Saving processing data to log
        SendInfoToLog(vbCrLf &
                          "            GetSonographs completed. Results" & vbCrLf &
                          "            Total number of words " & MemberWords.Count & vbCrLf &
                          "            Number of excluded abbreviations " & ExcludedAbbreviations & vbCrLf &
                          "            Number of excluded foreign words " & ExcludedForeignWords & vbCrLf &
                          "            Number of excluded InclusiveWordClassDetectionWords " & Excluded_InclusiveWordClassDetectionWords & vbCrLf &
                          "            Number of included InclusiveWordClassDetectionWords " & Included_InclusiveWordClassDetectionWords & vbCrLf &
                          "            Number of excluded ExclusiveWordClassDetectionWords " & Excluded_ExclusiveWordClassDetectionWords & vbCrLf &
                          "            Number of included ExclusiveWordClassDetectionWords " & Included_ExclusiveWordClassDetectionWords & vbCrLf &
                          "            Number of excluded ExclusiveWordClassCombinationDetectionWords " & Excluded_ExclusiveWordClassCombinationDetectionWords & vbCrLf &
                          "            Number of included ExclusiveWordClassCombinationDetectionWords " & Included_ExclusiveWordClassCombinationDetectionWords & vbCrLf &
                          "            Number of included/processed words " & ProcessedWordsCount)

        'Closing the progress display
        myProgress.Close()

        Return p2g_Settings.PhoneToGraphemesDictionary

    End Function


#End Region

#Region "Neighborhood"


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="CorpusTotalTokenCount"></param>
    ''' <param name="CorpusTotalWordTypeCount"></param>
    ''' <param name="PositionTerm"></param>
    ''' <param name="PrecalculatedZipfValues">Set to true if the values stored in the PLD1 transcriptions are Zipf values, or false if they are raw word frequencies.</param>
    ''' <param name="InactivateLog"></param>
    Public Sub Calculate_FWPN_DensityProbability(Optional ByVal CorpusTotalTokenCount As Long? = Nothing, Optional ByVal CorpusTotalWordTypeCount As Integer? = Nothing,
                                                      Optional ByVal PositionTerm As Integer = 3,
                                                     Optional PrecalculatedZipfValues As Boolean = False,
                                                     Optional InactivateLog As Boolean = False)

        If InactivateLog = False Then SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

        If PrecalculatedZipfValues = False Then
            'Setting corpus description data (using the data for the Current wordGroup if not set by the optional parameters.)
            'If zipf values are precalculated, this is not needed.
            If CorpusTotalTokenCount Is Nothing Then CorpusTotalTokenCount = Me.CorpusTokenCount
            If CorpusTotalWordTypeCount Is Nothing Then CorpusTotalWordTypeCount = Me.CorpusWordTypeCount
        End If


        For word = 0 To MemberWords.Count - 1

            'Putting zipf values of the source word and its neighbors (stored in MemberWords(word).PLD1Transcriptions ) in a new list
            Dim NeighbourZipfValues As New List(Of Double)
            For NeighbourIndex = 0 To MemberWords(word).PLD1Transcriptions.Count - 1

                Dim CurrentNeighbourSplit() As String

                'Allows for parsing of the previous delimiter (semicolon)
                If MemberWords(word).PLD1Transcriptions(NeighbourIndex).Contains(";") Then
                    'Parsing by using semicolon (old versions)
                    CurrentNeighbourSplit = MemberWords(word).PLD1Transcriptions(NeighbourIndex).Trim.Split(";")
                Else
                    'Parsing by using colon (current version)
                    CurrentNeighbourSplit = MemberWords(word).PLD1Transcriptions(NeighbourIndex).Trim.Split(":")
                End If

                If CurrentNeighbourSplit.Length > 1 Then

                    If IsNumeric(CurrentNeighbourSplit(1).Trim) Then

                        'Adding the neighbour value (as well as the source word, which is at the first index in PLD1Transcriptions)

                        If PrecalculatedZipfValues = False Then
                            'The frequency values stored in the transcription are raw frequencies. Calculates therefore their zipf values
                            Dim CurrentRawFrequency As Integer = CurrentNeighbourSplit(1).Trim
                            NeighbourZipfValues.Add(CalculateZipfValue(CurrentRawFrequency, CorpusTotalTokenCount, CorpusTotalWordTypeCount, PositionTerm))

                        Else
                            'The frequency values stored in the transcription are Zipf values. Adding them straight away
                            NeighbourZipfValues.Add(CurrentNeighbourSplit(1).Trim)
                        End If

                    Else
                        MsgBox("Error in " & System.Reflection.MethodInfo.GetCurrentMethod.Name & vbCrLf &
                               "Non numeric LD1-Neighbour frequency value: " & MemberWords(word).PLD1Transcriptions(NeighbourIndex))
                    End If
                Else
                    MsgBox("Error in " & System.Reflection.MethodInfo.GetCurrentMethod.Name & vbCrLf &
                               "LD1-Neighbour without word frequency data. " & MemberWords(word).PLD1Transcriptions(NeighbourIndex))
                End If

            Next

            'Getting the Zipf value sum of the group
            Dim ZipfSum As Double = 0
            For LDFrequency = 0 To NeighbourZipfValues.Count - 1
                ZipfSum += NeighbourZipfValues(LDFrequency)
            Next

            'Calculating density word probability
            If NeighbourZipfValues.Count > 0 Then
                MemberWords(word).FWPN_DensityProbability = NeighbourZipfValues(0) / ZipfSum
            Else
                MemberWords(word).FWPN_DensityProbability = 1
            End If

        Next

        If InactivateLog = False Then SendInfoToLog("   " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " completed successfully.")

    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="CorpusTotalTokenCount"></param>
    ''' <param name="CorpusTotalWordTypeCount"></param>
    ''' <param name="PositionTerm"></param>
    ''' <param name="PrecalculatedZipfValues">Set to true if the values stored in the OLD1 spellings are Zipf values, or false if they are raw word frequencies.</param>
    ''' <param name="InactivateLog"></param>
    Public Sub Calculate_FWON_DensityProbability(Optional ByVal CorpusTotalTokenCount As Long? = Nothing, Optional ByVal CorpusTotalWordTypeCount As Integer? = Nothing,
                                                      Optional ByVal PositionTerm As Integer = 3,
                                                     Optional PrecalculatedZipfValues As Boolean = False,
                                                     Optional InactivateLog As Boolean = False)

        If InactivateLog = False Then SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

        If PrecalculatedZipfValues = False Then
            'Setting corpus description data (using the data for the Current wordGroup if not set by the optional parameters.)
            'If zipf values are precalculated, this is not needed.
            If CorpusTotalTokenCount Is Nothing Then CorpusTotalTokenCount = Me.CorpusTokenCount
            If CorpusTotalWordTypeCount Is Nothing Then CorpusTotalWordTypeCount = Me.CorpusWordTypeCount
        End If


        For word = 0 To MemberWords.Count - 1

            'Putting zipf values of the source word and its neighbors (stored in MemberWords(word).OLD1Spellings ) in a new list
            Dim NeighbourZipfValues As New List(Of Double)
            For NeighbourIndex = 0 To MemberWords(word).OLD1Spellings.Count - 1

                Dim CurrentNeighbourSplit() As String = MemberWords(word).OLD1Spellings(NeighbourIndex).Trim.Split(",")

                If CurrentNeighbourSplit.Length > 1 Then

                    If IsNumeric(CurrentNeighbourSplit(1).Trim) Then

                        'Adding the neighbour value (as well as the source word, which is at the first index in OLD1Spellings)
                        If PrecalculatedZipfValues = False Then
                            'The frequency values stored in the transcription are raw frequencies. Calculates therefore their zipf values
                            Dim CurrentRawFrequency As Integer = CurrentNeighbourSplit(1).Trim
                            NeighbourZipfValues.Add(CalculateZipfValue(CurrentRawFrequency, CorpusTotalTokenCount, CorpusTotalWordTypeCount, PositionTerm))

                        Else
                            'The frequency values stored in the transcription are Zipf values. Adding them straight away
                            NeighbourZipfValues.Add(CurrentNeighbourSplit(1).Trim)
                        End If

                    Else
                        MsgBox("Error in " & System.Reflection.MethodInfo.GetCurrentMethod.Name & vbCrLf &
                               "Non numeric OLD1-Neighbour frequency value: " & MemberWords(word).OLD1Spellings(NeighbourIndex))
                    End If
                Else
                    MsgBox("Error in " & System.Reflection.MethodInfo.GetCurrentMethod.Name & vbCrLf &
                               "OLD1-Neighbour without word frequency data. " & MemberWords(word).OLD1Spellings(NeighbourIndex))
                End If

            Next

            'Getting the Zipf value sum of the group
            Dim ZipfSum As Double = 0
            For LDFrequency = 0 To NeighbourZipfValues.Count - 1
                ZipfSum += NeighbourZipfValues(LDFrequency)
            Next

            'Calculating density word probability
            If NeighbourZipfValues.Count > 0 Then
                MemberWords(word).FWON_DensityProbability = NeighbourZipfValues(0) / ZipfSum
            Else
                MemberWords(word).FWON_DensityProbability = 1
            End If

        Next

        If InactivateLog = False Then SendInfoToLog("   " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " completed successfully.")

    End Sub



    Public Sub Update_PLD1NeighbourCount()

        SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

        For Each MemberWord In MemberWords
            If MemberWord.PLD1Transcriptions Is Nothing Then
            Else
                If MemberWord.PLD1Transcriptions.Count = 0 Then
                    MemberWord.PLD1WordCount = 0
                Else
                    MemberWord.PLD1WordCount = MemberWord.PLD1Transcriptions.Count - 1 '-1 is used since the first transcription in the PLD1Transcriptions array is the current member word
                End If
            End If
        Next

        SendInfoToLog("   " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " completed successfully.")

    End Sub

    Public Sub Update_OLD1NeighbourCount()

        SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

        For Each MemberWord In MemberWords
            If MemberWord.OLD1Spellings Is Nothing Then
            Else
                If MemberWord.OLD1Spellings.Count = 0 Then
                    MemberWord.OLD1WordCount = 0
                Else
                    MemberWord.OLD1WordCount = MemberWord.OLD1Spellings.Count - 1 '-1 is used since the first transcription in the PLD1Transcriptions array is the current member word
                End If
            End If
        Next

        SendInfoToLog("   " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " completed successfully.")

    End Sub


    ''' <summary>
    ''' Holds phonetic comparison corpus data, consisting of transcriptions and their corresponding Zipf-value
    ''' </summary>
    Public Class PhoneticComparisonCorpus
        Inherits List(Of Tuple(Of List(Of String), Single))

    End Class


    ''' <summary>
    ''' Holds orthographic comparison corpus data, consisting of spellings and their corresponding Zipf-value
    ''' </summary>
    Public Class OrthographicComparisonCorpus
        Inherits Dictionary(Of String, Single) 'This was changed from List of Tuple to dictionary on 2019-08-16, to increase efficiency when adding new words to the comparison corpus on the website.
        'Inherits List(Of Tuple(Of String, Single))

        ''' <summary>
        ''' Loads a .txt file file containing OLDComparisonCorpus data, and returns a new OLDComparisonCorpus instance containing that data.
        ''' </summary>
        ''' <param name="filePath"></param>
        Public Shared Function LoadComparisonCorpus(Optional ByRef filePath As String = "", Optional ErrorString As String = "") As OrthographicComparisonCorpus

            Try

                Dim InputLines() As String

                'Getting the file path
                If filePath = "" Then
                    Dim dataString As String = My.Resources.OLDComparisonCorpus_ArcList
                    dataString = dataString.Replace(vbCrLf, vbLf)
                    InputLines = dataString.Split(vbLf)
                Else
                    InputLines = System.IO.File.ReadAllLines(filePath, Text.Encoding.UTF8)
                End If


                Dim output As New OrthographicComparisonCorpus

                For LineIndex = 1 To InputLines.Length - 1 'Skipping heading line

                    If InputLines(LineIndex).Trim = "" Then Continue For

                    'Reading all data in the current line
                    Dim LineSplit() As String = InputLines(LineIndex).Split(vbTab)
                    Dim CurrentSpelling As String = LineSplit(0).Trim
                    Dim CurrentZipfValue As Double = LineSplit(1).Trim

                    'Adding the current OLD spelling and frequency data
                    If Not output.ContainsKey(CurrentSpelling) Then
                        output.Add(CurrentSpelling, CurrentZipfValue)
                    Else
                        ErrorString &= "There are multiple identical spellings in the OLD comparison corpus file! This is not allowed."
                        Return Nothing
                    End If

                    'output.Add(New Tuple(Of String, Single)(CurrentSpelling, CurrentZipfValue))

                Next

                Return output

            Catch ex As Exception
                ErrorString &= ex.ToString
                Return Nothing
            End Try

        End Function


        Public Sub ExportToTxtFile(Optional ByRef saveDirectory As String = "", Optional ByRef saveFileName As String = "OLDComparisonCorpus",
                                            Optional BoxTitle As String = "Choose location to store the OLDComparisonCorpus output data...")

            Try

                SendInfoToLog("Attempting to export OLDComparisonCorpus data")

                'Choosing file location
                Dim filepath As String = ""
                'Ask the user for file path if not incomplete file path is given
                If saveDirectory = "" Or saveFileName = "" Then
                    filepath = GetSaveFilePath(saveDirectory, saveFileName, {"txt"}, BoxTitle)
                Else
                    filepath = Path.Combine(saveDirectory, saveFileName & ".txt")
                    If Not Directory.Exists(Path.GetDirectoryName(filepath)) Then Directory.CreateDirectory(Path.GetDirectoryName(filepath))
                End If


                'Saving to file
                Dim writer As New StreamWriter(filepath, False, Text.Encoding.UTF8)

                writer.WriteLine("Spelling" & vbTab & "Zipf-value")

                For Each CurrentWord In Me
                    writer.WriteLine(CurrentWord.Key & vbTab & CurrentWord.Value)
                Next

                writer.Close()

                SendInfoToLog("OLDComparisonCorpus data was exported to file: " & filepath)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

    End Class



    ''' <summary>
    ''' Holds lists of PLDComparisonCorpusData separated by syllable length
    ''' </summary>
    Public Class PLDComparisonCorpus
        Inherits SortedList(Of Integer, PLD_SyllableLengthSpecificComparisonCorpusData)


        ''' <summary>
        ''' Loads a .txt file file containing PLDComparisonCorpus data, and returns a new PLDComparisonCorpus instance containing that data.
        ''' </summary>
        ''' <param name="filePath"></param>
        Public Shared Function LoadComparisonCorpus(Optional ByRef filePath As String = "") As PLDComparisonCorpus

            Try

                Dim InputLines() As String

                'Getting the file path
                If filePath = "" Then
                    Dim dataString As String = My.Resources.PLDComparisonCorpus_ArcList
                    dataString = dataString.Replace(vbCrLf, vbLf)
                    InputLines = dataString.Split(vbLf)
                Else
                    InputLines = System.IO.File.ReadAllLines(filePath, Text.Encoding.UTF8)
                End If

                Dim output As New PLDComparisonCorpus

                For LineIndex = 1 To InputLines.Length - 1 'Skipping heading line

                    If InputLines(LineIndex).Trim = "" Then Continue For

                    'Reading all data in the current line
                    Dim LineSplit() As String = InputLines(LineIndex).Split(vbTab)

                    Dim CurrentSyllableLength As Integer = LineSplit(0)
                    Dim CurrentTranscriptionString() As String = LineSplit(1).Split(" ")
                    Dim CurrentTranscription As New List(Of String)
                    For Each TranscriptionItem In CurrentTranscriptionString
                        CurrentTranscription.Add(TranscriptionItem)
                    Next
                    Dim CurrentZipfValue As Double = LineSplit(2)

                    'Checking if the current syllable length is added
                    If Not output.ContainsKey(CurrentSyllableLength) Then output.Add(CurrentSyllableLength, New PLD_SyllableLengthSpecificComparisonCorpusData)

                    'Adding the current PLD1 transcription and frequency data

                    output(CurrentSyllableLength).Add(String.Join(" ", CurrentTranscription), New PLD_ComparisonCorpusData With {.PLD1Transcription = CurrentTranscription, .ZipfValue = CurrentZipfValue})
                    'output(CurrentSyllableLength).Add(New PLDComparisonCorpusData With {.PLD1Transcription = CurrentTranscription, .ZipfValue = CurrentZipfValue})

                Next

                Return output

            Catch ex As Exception
                Return Nothing
            End Try

        End Function


        Public Sub ExportToTxtFile(Optional ByRef saveDirectory As String = "", Optional ByRef saveFileName As String = "PLDComparisonCorpus",
                                            Optional BoxTitle As String = "Choose location to store the PLDComparisonCorpus output data...")

            Try

                SendInfoToLog("Attempting to export PLDComparisonCorpus data")

                'Choosing file location
                Dim filepath As String = ""
                'Ask the user for file path if not incomplete file path is given
                If saveDirectory = "" Or saveFileName = "" Then
                    filepath = GetSaveFilePath(saveDirectory, saveFileName, {"txt"}, BoxTitle)
                Else
                    filepath = Path.Combine(saveDirectory, saveFileName & ".txt")
                    If Not Directory.Exists(Path.GetDirectoryName(filepath)) Then Directory.CreateDirectory(Path.GetDirectoryName(filepath))
                End If


                'Saving to file
                Dim writer As New StreamWriter(filepath, False, Text.Encoding.UTF8)

                writer.WriteLine("Syllable length" & vbTab & "PLD1 transcription" & vbTab & "Zipf-value")

                For Each SyllableLength In Me
                    For Each PLD1Transcription In SyllableLength.Value

                        writer.WriteLine(SyllableLength.Key & vbTab & String.Join(" ", PLD1Transcription.Value.PLD1Transcription) &
                                             vbTab & PLD1Transcription.Value.ZipfValue)

                    Next
                Next

                writer.Close()

                SendInfoToLog("PLDComparisonCorpus data was exported to file: " & filepath)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

    End Class


    ''' <summary>
    ''' Holds a dictionary of PLDComparisonCorpusData of a certain syllable length
    ''' </summary>
    <Serializable>
    Public Class PLD_SyllableLengthSpecificComparisonCorpusData
        Inherits Dictionary(Of String, PLD_ComparisonCorpusData)
    End Class


    '' <summary>
    '' The PLDComparisonCorpusData of one single PLDComparison word
    '' </summary>
    <Serializable>
    Public Class PLD_ComparisonCorpusData
        Public PLD1Transcription As New List(Of String)
        Public ZipfValue As Single
    End Class


    ''' <summary>
    ''' Idenitifies PLD1 words of each memberword, based on the data in the ComparisonCorpus, optionally also caluculates PLDx.
    ''' Zipf valeus of each memberword must be calculated in advance and stored in the ZipfValue_Word property (use the CalculateZipfValue method!).
    ''' </summary>
    ''' <param name="ComparisonCorpus"></param>
    ''' <param name="MinimumWordCount">The number of included neighbors in the PLDx data. For example, set to 20 to calculate PLD20.</param>
    Public Sub Calculate_PLD_UsingComparisonCorpus(ByRef ComparisonCorpus As PLDComparisonCorpus,
                                                     Optional ByVal ComparisonCorpuTotalTokenCount As Long = 536866005,
                                                      Optional ByVal ComparisonCorpuTotalWordTypeCount As Integer = 3591552,
                                                       Optional ByVal ErrorString As String = "",
                                                       Optional ByVal CalculatePLDx As Boolean = True,
                                                       Optional ByVal MinimumWordCount As Integer = 20) 'Optional values based on the ARC-list word frequency data corpus size


        SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

        'Setting up a connection to the ARC-list database
        'Dim AfcListConnection = New MySql.Data.MySqlClient.MySqlConnection(AfcListMySqlConnectionString)

        'Recreating the PLD1 transcriptions, and put them in CurrentWord.TranscriptionString 
        For Each CurrentWord In MemberWords

            'Recreating the current word PLD1 transcription
            CurrentWord.Phonemes = CurrentWord.BuildPLD1TypeTranscription
            CurrentWord.TranscriptionString = String.Join(" ", CurrentWord.Phonemes)

            'Also resetting any values in the PLD1Transcriptions array
            CurrentWord.PLD1Transcriptions = New List(Of String)
        Next

        'Initializing PLDx data for each member word
        If CalculatePLDx = True Then
            For Each CurrentWord In MemberWords
                CurrentWord.PLDxData = New List(Of Tuple(Of Integer, String, Single))
            Next
        End If

        'Calculates PLD of all words using in the current word group, based on the comparison corpus
        For Each CurrentSyllableLength In ComparisonCorpus.Keys

            'Looking at each syllable length that exist in the comparison corpus
            For Each MemberWord In MemberWords

                'Only comparing if the words have the same syllable count
                If MemberWord.Syllables.Count = CurrentSyllableLength Then

                    'Dim SourceWordZipfValue? As Double = Nothing

                    For Each ComparisonWord In ComparisonCorpus(CurrentSyllableLength)

                        'Here it's possible to limit the number of comparisons made if calculation takes too long...
                        'Only comparing if the length of the words is +/-5 phoneme
                        'If Math.Abs(MemberWord.Phonemes.Count - ComparisonWord.Value.PLD1Transcription.Count) > 10 Then Continue For

                        'Only comparing if the length of the words is +/-1 phoneme
                        'If MemberWord.Phonemes.Count = ComparisonWord.Value.PLD1Transcription.Count Or
                        '            MemberWord.Phonemes.Count = ComparisonWord.Value.PLD1Transcription.Count - 1 Or
                        '            MemberWord.Phonemes.Count = ComparisonWord.Value.PLD1Transcription.Count + 1 Then


                        'Calculates PLD1
                        Dim CurrentPLD As Integer = LevenshteinDistance(MemberWord.Phonemes, ComparisonWord.Value.PLD1Transcription)

                        'Adding the current PLD1 word if its transcription part is not already added 
                        Dim ComparisonWordPLD1String As String = String.Join(" ", ComparisonWord.Value.PLD1Transcription)

                        If CurrentPLD = 0 Then 'I.e. It is the same word
                            'Setting the Zipf value of the current word, if it's included in the ComparisonCorpus
                            'SourceWordZipfValue = ComparisonWord.Value.ZipfValue

                        ElseIf CurrentPLD = 1 Then 'It's a PLD1 neighbor

                            'Putting all existing PLD1 transcriptions for the current MemberWord in a SortedSet 
                            Dim TempTranscriptionPartList As New SortedSet(Of String)
                            For Each ExistingPLD1Transcription In MemberWord.PLD1Transcriptions
                                TempTranscriptionPartList.Add(ExistingPLD1Transcription.Split(":")(0))
                            Next

                            'Adding the currently detected PLD1 neighbor only if it's not already added as a neighbor of the current MemberWord
                            If Not TempTranscriptionPartList.Contains(ComparisonWordPLD1String) Then
                                MemberWord.PLD1Transcriptions.Add(ComparisonWordPLD1String & ":" & ComparisonWord.Value.ZipfValue)
                            End If

                        End If

                        'Storing PLDx words
                        If CalculatePLDx = True Then
                            If CurrentPLD > 0 Then
                                MemberWord.PLDxData.Add(New Tuple(Of Integer, String, Single)(CurrentPLD, ComparisonWordPLD1String, ComparisonWord.Value.ZipfValue))

                                'Sorting the words in rising order according to LD value, and then in falling order after Zipf-scale value
                                Dim SortedPldxData As New List(Of Tuple(Of Integer, String, Single))
                                Dim Query = MemberWord.PLDxData.OrderBy(Function(myTuple) myTuple.Item1).ThenByDescending(Function(myTuple) myTuple.Item3)
                                For Each CurrentTuple In Query
                                    SortedPldxData.Add(CurrentTuple)
                                Next
                                MemberWord.PLDxData = SortedPldxData

                                'Removes the last LD word if it is outside the desired minimum LD words count
                                If MemberWord.PLDxData.Count > MinimumWordCount Then
                                    MemberWord.PLDxData.RemoveAt(MemberWord.PLDxData.Count - 1)
                                End If

                            End If
                        End If
                        'End If
                    Next

                    'Sorting the words in MemberWord.PLD1Transcriptions in falling order after word frequency (Zipf-scale value)
                    MemberWord.SortPLD1Transcriptions(False) 'False is set here, since the current MemberWord is not yet added (Adding it later, should make the sorting a little faster.)

                    'Adding the value of the current MemberWord to it's own PLD1Transcriptions at index 0, only if the word has at least one neighbor.
                    If MemberWord.PLD1Transcriptions.Count > 0 Then MemberWord.PLD1Transcriptions.Insert(0, String.Join(" ", MemberWord.Phonemes) & ":" & MemberWord.ZipfValue_Word)

                End If
            Next
        Next

        'Averaging PLDx data
        If CalculatePLDx = True Then

            For Each MemberWord In MemberWords

                'Summing PLDx values 
                Dim SummedValueList As New List(Of Single)
                For Each DataPoint In MemberWord.PLDxData
                    SummedValueList.Add(DataPoint.Item1)
                Next

                'getting and storing the average (or resetting the default value of 0 if no PLDx words exist)
                If SummedValueList.Count > 0 Then
                    MemberWord.PLDx_Average = SummedValueList.Average
                Else
                    MemberWord.PLDx_Average = 0
                End If

                'Inserting the source word PLD transcription into index 0 for reference, and consistency with the PLD1 transcriptions
                If MemberWord.PLDxData.Count > 1 Then
                    MemberWord.PLDxData.Insert(0, New Tuple(Of Integer, String, Single)(0, String.Join(" ", MemberWord.Phonemes), MemberWord.ZipfValue_Word))
                End If

            Next

        End If

        SendInfoToLog("Method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " completed successfully.")

    End Sub


    ''' <summary>
    ''' Calculates OLD1 and optionally OLDx based on the words in the comparison corpus.
    ''' </summary>
    ''' <param name="ComparisonCorpus"></param>
    ''' <param name="ErrorString"></param>
    ''' <param name="CalculateOLDx">Set to false to skip calculation of OLDx</param>
    ''' <param name="MinimumWordCount">The number of words to be included when calculating mean OLD of the MinimumWordCount closest neighbours (E.G. use 20 to calculate OLD20).</param>
    Public Sub Calculate_OLD_UsingComparisonCorpus(ByRef ComparisonCorpus As OrthographicComparisonCorpus,
                                                   Optional ByVal ErrorString As String = "",
                                                   Optional ByVal CalculateOLDx As Boolean = True,
                                                   Optional ByVal MinimumWordCount As Integer = 20,
                                                   Optional ByVal IsRunOnServer As Boolean = False)


        'The ComparisonCorpus, should contain [Spelling, Zipf-scale value]

        'Logging initialization
        If IsRunOnServer = False Then SendInfoToLog("Initializing calculation of LevenshteinDistances of type: ")

        Dim ProgressForm As ProgressDisplay = Nothing
        If IsRunOnServer = False Then
            ProgressForm = New ProgressDisplay
            ProgressForm.Initialize(MemberWords.Count - 1,, "Calculating OLD")
            ProgressForm.Show()
        End If

        Try

            For SourceWordIndex = 0 To MemberWords.Count - 1

                'Updating progress
                If IsRunOnServer = False Then ProgressForm.UpdateProgress(SourceWordIndex)

                Dim CurrentWord As Word = MemberWords(SourceWordIndex)

                'Resetting any values in the OLD1Spellings array
                CurrentWord.OLD1Spellings = New List(Of String)

                'Also resetting the OLDxData of the word
                If CalculateOLDx = True Then CurrentWord.OLDxData = New List(Of Tuple(Of Integer, String, Single))

                'Calculating OLD
                For Each ComparisonWord In ComparisonCorpus

                    Dim CurrentLD As Integer = LevenshteinDistance(ComparisonWord.Key, CurrentWord.OrthographicForm)

                    'Storing OLD1 data
                    If CurrentLD < 2 Then
                        CurrentWord.OLD1Spellings.Add(String.Concat(ComparisonWord.Key) & "," & ComparisonWord.Value)
                        'Comma is used instead of colon since is a part of some spellings.)
                        'N.B. This function exports Zipf-scale value, instead of the raw word frequency value in the AFC-list column OLD1Spellings! The reason is that the comparison corpus contains Zipf scale values.
                    End If

                    'Storing OLDx data
                    If CalculateOLDx = True Then

                        'Skipping if it's the same word
                        If CurrentLD <= 0 Then Continue For '(If CurrentLD = 0 it is the same wordtype. CurrentLD should never be below 0.)

                        CurrentWord.OLDxData.Add(New Tuple(Of Integer, String, Single)(CurrentLD, ComparisonWord.Key, ComparisonWord.Value))

                        'Sorting the words in rising order according to LD value
                        'Sorting the words in rising order according to LD value, and then in falling order after Zipf-scale value
                        Dim SortedOldxData As New List(Of Tuple(Of Integer, String, Single))
                        Dim Query = CurrentWord.OLDxData.OrderBy(Function(myTuple) myTuple.Item1).ThenByDescending(Function(myTuple) myTuple.Item3)
                        For Each CurrentTuple In Query
                            SortedOldxData.Add(CurrentTuple)
                        Next
                        CurrentWord.OLDxData = SortedOldxData

                        'Removes the last LD word if it is outside the desired minimum LD words count
                        If CurrentWord.OLDxData.Count > MinimumWordCount Then
                            CurrentWord.OLDxData.RemoveAt(CurrentWord.OLDxData.Count - 1)
                        End If

                    End If
                Next

                'Removing the OLD1Spellings of the current word if only the source word was detected
                If CurrentWord.OLD1Spellings.Count = 1 Then CurrentWord.OLD1Spellings = New List(Of String)


                If CalculateOLDx = True Then
                    'Summing OLDx values 
                    Dim SummedValueList As New List(Of Single)
                    For Each DataPoint In MemberWords(SourceWordIndex).OLDxData
                        SummedValueList.Add(DataPoint.Item1)
                    Next

                    'getting and storing the average
                    MemberWords(SourceWordIndex).OLDx_Average = SummedValueList.Average

                    'Inserting the source word PLD transcription into index 0 for reference, and consistency with the PLD1 transcriptions
                    If CurrentWord.OLDxData.Count > 1 Then
                        CurrentWord.OLDxData.Insert(0, New Tuple(Of Integer, String, Single)(0, CurrentWord.OrthographicForm, CurrentWord.ZipfValue_Word))
                    End If
                End If

            Next

            If IsRunOnServer = False Then ProgressForm.Hide()

            'Logging
            If IsRunOnServer = False Then SendInfoToLog("Completed calculation of OLD")

        Catch ex As Exception
            ErrorString &= ex.ToString
            Exit Sub
        End Try

    End Sub



    <Serializable>
    Public Class LD_String
        Implements IComparable

        Public Sub New()

        End Sub
        'Implements IComparable(Of String)

        Public Property LevenshteinDifference As Integer
        Public Property Word As Word
        'Public Property OrthographicForm As String
        'Public Property Phonemes As New List(Of String)

        'Public Function CompareTo(other As String) As Integer Implements IComparable(Of String).CompareTo
        'Return OrthographicForm.CompareTo(other)
        'End Function

        Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            If Not TypeOf (obj) Is LD_String Then
                Throw New ArgumentException()
            Else

                Dim tempOLD As LD_String = DirectCast(obj, LD_String)

                If Me.LevenshteinDifference < tempOLD.LevenshteinDifference Then
                    Return -1
                ElseIf Me.LevenshteinDifference = tempOLD.LevenshteinDifference Then
                    Return 0
                Else
                    Return 1
                End If
            End If

        End Function

    End Class





#End Region


#Region "IsolationPoints"



    Public Sub CalculatePhoneticIsolationPoints(ByRef ComparisonCorpus As PLDComparisonCorpus,
                                                       Optional ByVal ErrorString As String = "",
                                                Optional ShowProgressWindow As Boolean = False)

        Try

            SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            If ComparisonCorpus Is Nothing Then
                'Create a comparison corpus, from the word in the wordgroup
                Throw New NotImplementedException
            End If

            Dim PhoneticIsolationPointCalculator As New PhoneticIsolationPoints(ComparisonCorpus)

            'Starting a progress window
            Dim myProgress As ProgressDisplay = Nothing
            If ShowProgressWindow = True Then
                myProgress = New ProgressDisplay
                myProgress.Initialize(MemberWords.Count - 1, 0, "Calculating phonetic isolation points...", 1)
                myProgress.Show()
            End If

            Dim CalculatedWordsCount As Integer = 0
            For WordIndex = 0 To MemberWords.Count - 1

                'Updating progress
                If ShowProgressWindow = True Then myProgress.UpdateProgress(WordIndex)

                Dim PLD1Transcription = MemberWords(WordIndex).BuildPLD1TypeTranscription
                'Calculating PhoneticIsolationPoint only if the PLD1Transcription is longer than 2 segments (as the initial stressed syllable index and tone data is removed by the method called)
                If PLD1Transcription.Count > 2 Then
                    MemberWords(WordIndex).PhoneticIsolationPoint = PhoneticIsolationPointCalculator.GetIsolationPoint(MemberWords(WordIndex).BuildPLD1TypeTranscription, ErrorString)
                    CalculatedWordsCount += 1

                    If ErrorString <> "" Then Errors(ErrorString)

                End If
            Next

            'Closing the progress display
            If ShowProgressWindow = True Then myProgress.Close()

            SendInfoToLog("Method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " completed successfully. Results: Phonetic isolation points were calculated for " & CalculatedWordsCount & " out of " & MemberWords.Count & " words.")
        Catch ex As Exception
            ErrorString &= ErrorString & vbCrLf
        End Try

    End Sub

    Public Sub CalculateOrthographicIsolationPoints(ByRef ComparisonCorpus As OrthographicComparisonCorpus,
                                                       Optional ByVal ErrorString As String = "",
                                                    Optional ByVal IsRunOnServer As Boolean = False)

        Try

            SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            If ComparisonCorpus Is Nothing Then
                'Create a comparison corpus, from the word in the wordgroup
                Throw New NotImplementedException
            End If

            Dim OrthographicIsolationPointCalculator As New OrthographicIsolationPoints(ComparisonCorpus)


            'Starting a progress window
            Dim myProgress As ProgressDisplay = Nothing
            If IsRunOnServer = False Then
                myProgress = New ProgressDisplay
                myProgress.Initialize(MemberWords.Count - 1, 0, "Calculating phonetic isolation points...", 1)
                myProgress.Show()
            End If

            Dim CalculatedWordsCount As Integer = 0
            For WordIndex = 0 To MemberWords.Count - 1

                'Updating progress
                If IsRunOnServer = False Then myProgress.UpdateProgress(WordIndex)

                MemberWords(WordIndex).OrthographicIsolationPoint = OrthographicIsolationPointCalculator.GetIsolationPoint(MemberWords(WordIndex).OrthographicForm, ErrorString)

                If ErrorString <> "" Then Errors(ErrorString)

            Next

            'Closing the progress display
            myProgress.Close()

            If IsRunOnServer = False Then SendInfoToLog("Method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " completed successfully. Results: Orthographic isolation points were calculated for all " & MemberWords.Count & " words.")

        Catch ex As Exception
            ErrorString &= ErrorString & vbCrLf
        End Try

    End Sub



#End Region




    Public Enum WordFrequencyUnit
        RawFrequency
        Log10RawFrequency ' Sum the log 10 (raw frequency) value of each occurence
        Log10RawFrequencySum 'Takes the log 10 of the sum of all raw frequencies
        NormalizedFrequency
        ZipfValue
        Log10ZipfValue
        WordType ' = word frequency = 1 for all words occuring in the word list (No token based weighting)
    End Enum




#Region "SelectionAndFilterring"

    ''' <summary>
    ''' Selects a number of random words from the current word group and stores in a new Wordgroup (The words in the new WordGroup are the same instances as in the origial group.)
    ''' </summary>
    ''' <param name="WordCount">The number of words to select. (All words will be returned if the number is higher that the member word count.)</param>
    ''' <returns></returns>
    Public Function SelectRandomWords(ByVal WordCount As Integer) As WordGroup

        Dim SampleList As New WordGroup
        SampleList.GetCorpusInfoFromOtherWordgroup(Me)

        If MemberWords.Count > WordCount Then
            'Adding a random selection of words

            SampleList.GetCorpusInfoFromOtherWordgroup(Me) 'Copying corpus description data
            Dim SamplePoints As New SortedSet(Of Integer)
            For n = 0 To WordCount - 1
                SamplePoints.Add(Int((n / WordCount) * MemberWords.Count))
            Next
            For n = 0 To Me.MemberWords.Count - 1
                If SamplePoints.Contains(n) Then SampleList.MemberWords.Add(MemberWords(n))
            Next

        Else
            'Adding all words
            For n = 0 To Me.MemberWords.Count - 1
                SampleList.MemberWords.Add(MemberWords(n))
            Next

        End If

        Return SampleList

    End Function


#End Region


#Region "InputOutput"


    Public Enum TxtFileOutputTypes
        FullPhonetic
    End Enum



    ''' <summary>
    ''' Stores the current instance of WordGroup to a txt file, with one word on each row
    ''' </summary>
    ''' <param name="saveDirectory">The directory where the file is saved.</param>
    ''' <param name="saveFileName">The filename the file to save.</param>
    ''' <param name="BoxTitle"></param>
    ''' <param name="OutputType"></param>
    ''' <param name="OutputWordClassFrequencies">Set this to true if the word contains word class frequency data, that should be save to file.</param>
    ''' <param name="DataOutputFilePathUsed">Upon return of the function, this variable holds the actual fie path used."></param>
    ''' <returns>Returns True if the save procedure completed, and False is saving failed.</returns>
    Public Function SaveToTxtFile(Optional ByVal saveDirectory As String = "",
                                          Optional ByVal saveFileName As String = "",
                                          Optional ByVal BoxTitle As String = "",
                                          Optional OutputType As TxtFileOutputTypes = TxtFileOutputTypes.FullPhonetic,
                                          Optional OutputWordClassFrequencies As Boolean = True,
                                          Optional OutputGraphemes As Boolean = True,
                                          Optional LogOutputFolder As String = "",
                                          Optional ByRef DataOutputFilePathUsed As String = "",
                                          Optional ByVal ColumnOrder As Object = Nothing,
                                          Optional OutputManualEvaluationsForForeignWords As Boolean = True,
                                          Optional OutputAmbigousSyllableBoundaryMarkers As Boolean = False,
                                          Optional OutputSampaCountToLog As Boolean = True,
                                          Optional OutputSampleFile As Boolean = True,
                                          Optional SampleFileLineCount As Integer = 1000,
                                          Optional RoundingDecimals As Integer = 4,
                                  Optional SkipCorpusDescriptionLine As Boolean = False) As Boolean

0:
        Dim SaveAttempts As Integer = 0

        Try

            'Setting a default column order
            If ColumnOrder Is Nothing Then
                ColumnOrder = New PhoneticTxtStringColumnIndices
                ColumnOrder.SetDefaultOrder()
            End If

            SendInfoToLog("Attempting to save to txt file...",, LogOutputFolder)

            'Dim filepath As String = ""
            'Ask the user for file path if not incomplete file path is given
            If saveDirectory = "" Or saveFileName = "" Then
                DataOutputFilePathUsed = GetSaveFilePath(saveDirectory, saveFileName, {"txt"}, BoxTitle)
            Else
                DataOutputFilePathUsed = Path.Combine(saveDirectory, saveFileName & ".txt")
                If Not Directory.Exists(Path.GetDirectoryName(DataOutputFilePathUsed)) Then Directory.CreateDirectory(Path.GetDirectoryName(DataOutputFilePathUsed))
            End If

            If OutputSampleFile = True And MemberWords.Count > SampleFileLineCount Then
                Dim SampleList As WordGroup = SelectRandomWords(SampleFileLineCount)

                SampleList.SaveToTxtFile(saveDirectory, saveFileName & "_SampleFile",, OutputType, OutputWordClassFrequencies, OutputGraphemes,
                                                 LogOutputFolder,, ColumnOrder, OutputManualEvaluationsForForeignWords, OutputAmbigousSyllableBoundaryMarkers,
                                                 OutputSampaCountToLog, False)
            End If





            'Tests which type of object ColumnOrder is
            'If it is the new type
            If TypeOf (ColumnOrder) Is PhoneticTxtStringColumnIndices Then

                Dim CurrentColumnOrder As PhoneticTxtStringColumnIndices = DirectCast(ColumnOrder, PhoneticTxtStringColumnIndices)

                Dim writer As New StreamWriter(DataOutputFilePathUsed, False, Text.Encoding.UTF8)

                Select Case OutputType
                    Case TxtFileOutputTypes.FullPhonetic
                        'Writing current word index (used for retreiving the start position of longer manual processes)
                        'and also column headers

                        Dim ColumnToWrite As Integer = 0

                        If SkipCorpusDescriptionLine = False Then
                            'Writing settings / corpus descriptions on the second line
                            Dim SettingsLine As String = ""
                            SettingsLine &= "CorpusTokenCount=" & CorpusTokenCount & vbTab
                            SettingsLine &= "CorpusWordTypeCount=" & CorpusWordTypeCount & vbTab
                            SettingsLine &= "CorpusSentenceCount=" & CorpusSentenceCount & vbTab
                            SettingsLine &= "CorpusDocumentCount=" & CorpusDocumentCount & vbTab
                            SettingsLine &= "CurrentWordIndex=" & CurrentWordIndex & vbTab
                            writer.WriteLine(SettingsLine)
                        End If

                        'Writing headings on the second line
                        Dim HeadingLine As String = ""
                        'Counting the number of columns to export
                        Dim ColumnOrderProperyInfo() As System.Reflection.PropertyInfo = GetType(PhoneticTxtStringColumnIndices).GetProperties
                        For n = 0 To ColumnOrderProperyInfo.Length - 1
                            If GetType(PhoneticTxtStringColumnIndices).GetProperty(ColumnOrderProperyInfo(n).Name).GetValue(CurrentColumnOrder) IsNot Nothing Then
                                If GetType(PhoneticTxtStringColumnIndices).GetProperty(ColumnOrderProperyInfo(n).Name).GetValue(CurrentColumnOrder) = ColumnToWrite Then
                                    HeadingLine &= ColumnOrderProperyInfo(n).Name & vbTab
                                    ColumnToWrite += 1
                                End If
                            End If
                        Next

                        writer.WriteLine(HeadingLine)

                        'Writing word data

                        'Starting a progress window
                        Dim myProgress As New ProgressDisplay
                        myProgress.Initialize(MemberWords.Count - 1, 0, "Saving to file...", 100)
                        myProgress.Show()

                        For word = 0 To Me.MemberWords.Count - 1

                            'Updating progress
                            myProgress.UpdateProgress(word)

                            writer.WriteLine(MemberWords(word).GenerateFullPhoneticOutputTxtString(CurrentColumnOrder, RoundingDecimals))
                        Next

                        'Closing the progress display
                        myProgress.Close()

                        'SendInfoToLog("Total number of Sampa forms included on export: " & TotalNumberOfSampaForms,, LogOutputFolder)

                End Select

                writer.Close()

                If OutputSampaCountToLog = True Then
                    SendInfoToLog("     " & MemberWords.Count & " words have been saved to the file " & DataOutputFilePathUsed,, LogOutputFolder)
                End If

            End If



            Return True
        Catch ex As Exception
            MsgBox(ex.ToString)

            If SaveAttempts < 3 Then
                SaveAttempts += 1
                MsgBox("Please close any open instances of the file and press Ok to try to save again..." & vbCr & "Attempt " & SaveAttempts & " of 3.")
                SendInfoToLog("Save failed. Attempts again. Number " & SaveAttempts & " of 3 attempts.")
                GoTo 0
            Else
                SendInfoToLog("Save failed after 3 attempts. Exiting sub without saving.")
                Return False
            End If

        End Try
    End Function


    ''' <summary>
    ''' Loads an txt-file created using GeneratePhgoneticOutputTxtString (.txt) file containing phonetic info on a orthographic word, and returns a new Word containing that data.
    ''' </summary>
    ''' <param name="filePath">The path to the file to be loaded.</param>
    ''' <param name="ReplaceXXbyYYList">Allowes for basic string replacements in the input (Word) lines. Replaces each Tuple.Item1 by a Tuple.Item2.</param>
    ''' <param name="InputHeadingConversions">An optinal SortedList by which enables backward compatibility if column names have been changed. Keys repressent the old names, and values repressent the column name in the presently used ColumnOrder.</param>
    ''' <returns>Returns a new Word containing the data in the chosen file.</returns>
    Public Shared Function LoadWordGroupPhoneticTxtFile(Optional ByRef filePath As String = "",
                                                     Optional ByVal ReadAbbreviationColumn As Boolean = True,
                                                     Optional ByVal ReadForeignWordColumn As Boolean = True,
                                                     Optional ByVal CorrectedSpellingColumnExists As Boolean = True,
                                                     Optional ByVal OutputSampaCountToLog As Boolean = True,
                                                     Optional ByRef ColumnOrder As Object = Nothing,
                                                     Optional ByVal ReadRegularSpellingColumn As Boolean = True,
                                                     Optional ByVal ReadGraphemesColumn As Boolean = True,
                                                     Optional ByRef ReadSupraSegmentalsFromTranscription As Boolean = True,
                                                     Optional ByRef CorrectDoubleSpacesInPhoneticForm As Boolean = True,
                                                     Optional ByRef CheckPhonemeValidity As Boolean = True,
                                                     Optional ByRef ValidPhoneticCharacters As List(Of String) = Nothing,
                                                     Optional ByVal DoSyllabicAnalysis As Boolean = True,
                                                     Optional ByRef ReplaceXXbyYYList As List(Of Tuple(Of String, String)) = Nothing,
                                                     Optional ByRef InputHeadingConversions As SortedList(Of String, String) = Nothing,
                                                     Optional ByRef ColumnLengthList As SortedList(Of String, Integer) = Nothing) As WordGroup

        'TODO: The column orders can be read automatically also for the word list editing column type, and thereby also use the same parsing function.
        'But a problem is that the first column is "high-jacked" by CurrentWord instead of OrthographicForm

        SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

        Try

            'Sets a default column order
            If ColumnOrder Is Nothing Then
                ColumnOrder = New PhoneticTxtStringColumnIndices
                ColumnOrder.SetDefaultOrder()
            End If


            Dim OutputWordGroup As New WordGroup

            'Tests which type of object ColumnOrder is
            'If it is the new type
            If TypeOf (ColumnOrder) Is PhoneticTxtStringColumnIndices Then

                Dim CurrentColumnOrder As PhoneticTxtStringColumnIndices = DirectCast(ColumnOrder, PhoneticTxtStringColumnIndices)

                If filePath = "" Then filePath = GetOpenFilePath(,, {"txt"}, "Please open a word list file...")


                Dim inputArray() As String = {}
                inputArray = System.IO.File.ReadAllLines(filePath)


                'Reading the "CurrentWordIndex=" & CurrentWordIndex, and settings
                Dim FirstLineSplit() As String = inputArray(0).Split(vbTab)
                Dim DetectedCorpusDescriptionLine As Boolean = False

                Try
                    'Try Reading corpus description
                    Dim ProblemsDetected As Integer = 0
                    If FirstLineSplit(0).Split("=")(0).Trim.EndsWith("CorpusTokenCount") Then
                        OutputWordGroup.CorpusTokenCount = FirstLineSplit(0).Split("=")(1).Trim
                        DetectedCorpusDescriptionLine = True
                    Else
                        ProblemsDetected += 1
                    End If
                    If FirstLineSplit(1).Split("=")(0).Trim.EndsWith("CorpusWordTypeCount") Then
                        OutputWordGroup.CorpusWordTypeCount = FirstLineSplit(1).Split("=")(1).Trim
                        DetectedCorpusDescriptionLine = True
                    Else
                        ProblemsDetected += 1
                    End If
                    If FirstLineSplit(2).Split("=")(0).Trim.EndsWith("CorpusSentenceCount") Then
                        OutputWordGroup.CorpusSentenceCount = FirstLineSplit(2).Split("=")(1).Trim
                        DetectedCorpusDescriptionLine = True
                    Else
                        ProblemsDetected += 1
                    End If
                    If FirstLineSplit(3).Split("=")(0).Trim.EndsWith("CorpusDocumentCount") Then
                        OutputWordGroup.CorpusDocumentCount = FirstLineSplit(3).Split("=")(1).Trim
                        DetectedCorpusDescriptionLine = True
                    Else
                        ProblemsDetected += 1
                    End If

                    If ProblemsDetected > 0 Then
                        MsgBox("Detected problems reading CurrentWordIndex from line 1. Probably they are lacking or, incorrectly typed. Click ok to ignore corpus data, and assume that column headings are on line 1.")
                        DetectedCorpusDescriptionLine = False
                        OutputWordGroup.CorpusTokenCount = 0
                        OutputWordGroup.CorpusWordTypeCount = 0
                        OutputWordGroup.CorpusSentenceCount = 0
                        OutputWordGroup.CorpusDocumentCount = 0
                    End If

                Catch ex As Exception
                    MsgBox("Detected problems reading CurrentWordIndex from line 1. Probably they are lacking or, incorrectly typed. Click ok to ignore corpus data, and assume that column headings are on line 1.")
                    DetectedCorpusDescriptionLine = False
                    OutputWordGroup.CorpusTokenCount = 0
                    OutputWordGroup.CorpusWordTypeCount = 0
                    OutputWordGroup.CorpusSentenceCount = 0
                    OutputWordGroup.CorpusDocumentCount = 0
                End Try

                If FirstLineSplit.Count > 3 Then
                    Try
                        If FirstLineSplit(4).Split("=")(0).Trim = "CurrentWordIndex" Then
                            OutputWordGroup.CurrentWordIndex = FirstLineSplit(4).Split("=")(1).Trim
                            DetectedCorpusDescriptionLine = True
                        End If
                    Catch ex As Exception
                        MsgBox("Unable to read the current-word index. Click Ok to continue!")
                    End Try
                End If

                'Constructing a column order from the headings on the second line
                Dim HeadingLineIndex As Integer
                If DetectedCorpusDescriptionLine = False Then
                    HeadingLineIndex = 0
                Else
                    HeadingLineIndex = 1
                End If

                'Skipping a maximum of 10 empty lines before the headings line
                Dim StartSearchForHeadingIndex As Integer = HeadingLineIndex
                For StartSearchForHeadingIndex = 0 To StartSearchForHeadingIndex + 9
                    If inputArray(StartSearchForHeadingIndex).Trim = "" Then
                        HeadingLineIndex += 1
                    Else
                        'Found something on the line, assuming it to be headings
                        Exit For
                    End If
                Next

                Dim HeadingsLineSplit() As String = inputArray(HeadingLineIndex).Trim.Split(vbTab)
                Dim HeadingColumnOrder As New PhoneticTxtStringColumnIndices
                Dim Props() As PropertyInfo = GetType(PhoneticTxtStringColumnIndices).GetProperties
                Dim PropDict As New Dictionary(Of String, PropertyInfo)

                'Adding properties to the PropDict
                For i = 0 To Props.Length - 1
                    PropDict.Add(Props(i).Name, Props(i))
                Next

                If InputHeadingConversions Is Nothing Then InputHeadingConversions = New SortedList(Of String, String) From {}

                'Setting the values for the current column indices
                For CurrentHeading = 0 To HeadingsLineSplit.Length - 1
                    If PropDict.ContainsKey(HeadingsLineSplit(CurrentHeading).Trim) Then

                        'Using the standard column names
                        PropDict(HeadingsLineSplit(CurrentHeading).Trim).SetValue(HeadingColumnOrder, CurrentHeading)
                    ElseIf InputHeadingConversions.ContainsKey(HeadingsLineSplit(CurrentHeading).Trim) Then

                        'Using column name conversion
                        PropDict(InputHeadingConversions(HeadingsLineSplit(CurrentHeading).Trim)).SetValue(HeadingColumnOrder, CurrentHeading)

                    Else
                        MsgBox("Input word list contains an invalid column heading: " & HeadingsLineSplit(CurrentHeading) & vbCrLf &
                                   "That column will not be read.")
                        SendInfoToLog("Invalid input file column heading detected: " & HeadingsLineSplit(CurrentHeading) & "This column will not be read.")
                    End If
                Next

                'Adding keys to the ColumnLengthList 
                If ColumnLengthList IsNot Nothing Then
                    Dim Headings = HeadingColumnOrder.GetColumnHeadingsString().ToString.Split(vbTab)
                    For Each Item In Headings
                        If Item <> "" Then ColumnLengthList.Add(Item, 0)
                    Next
                End If

                Dim TotalSampaFormCount As Integer = 0

                'Starting a progress window
                Dim myProgress As New ProgressDisplay
                myProgress.Initialize(inputArray.Length - 1, 0, "Parsing input file...", 100)
                myProgress.Show()

                'Reading word data
                For i = HeadingLineIndex + 1 To inputArray.Count - 1

                    'Updating progress
                    myProgress.UpdateProgress(i)

                    'Skipping to next if the line is empty
                    If inputArray(i).Trim = "" Then Continue For

                    Dim InputLine As String = inputArray(i)

                    'Doing input string replacements
                    If ReplaceXXbyYYList IsNot Nothing Then
                        For r = 0 To ReplaceXXbyYYList.Count - 1
                            InputLine = InputLine.Replace(ReplaceXXbyYYList(r).Item1, ReplaceXXbyYYList(r).Item2)
                        Next
                    End If

                    OutputWordGroup.MemberWords.Add(Word.ParseInputWordString(InputLine, HeadingColumnOrder,
                                                                                            CorrectDoubleSpacesInPhoneticForm,
                                                                                            CheckPhonemeValidity, ValidPhoneticCharacters, ColumnLengthList))
                Next

                'Closing the progress display
                myProgress.Close()

            End If


            If DoSyllabicAnalysis = True Then
                'Analysing the input word list
                OutputWordGroup.DetermineSyllableIndices() 'Determining internal syllable structure
                OutputWordGroup.DetermineSyllableOpenness()  'Detecting syllable openness
            End If


            SendInfoToLog(OutputWordGroup.MemberWords.Count & " words were read from file " & filePath)

            Return OutputWordGroup
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return Nothing
        End Try

    End Function


#End Region

End Class


''' <summary>
''' Class that holds, and may compute, various information about a word.
''' </summary>
<Serializable>
Public Class Word

#Region "Declarations"
    'Variables that are read from the txt input file

    Public Property OrthographicForm As String
    Public Property GIL2P_OT_Average As Double
    Public Property GIL2P_OT_Min As Double
    Public Property PIP2G_OT_Average As Double
    Public Property PIP2G_OT_Min As Double
    Public Property G2P_OT_Average As Double
    Public Property ProportionStartingWithUpperCase As Double 'Percentage of times the word starts with a capital letter 
    Public Property LanguageHomographs As List(Of String)
    Public Property LanguageHomographCount As Single
    Public Property OrthographicFormContainsSpecialCharacter As Boolean = False
    Public Property RawWordTypeFrequency As Integer
    Public Property RawDocumentCount As Integer 'The number of document the specific word type exists in
    Public Property Syllables As New ListOfSyllables 'Stores the complete phonetic transcription segmented into syllables
    Public Property Syllables_AlternateSyllabification As ListOfSyllables 'Stores alternative syllabification, and can be used with BuildExtendedIpaArray_AlternateSyll
    Public Property PhonotacticType As String
    Public Property SSPP_Average As Double 'stress and syllable structure based normalized phonotactic probability
    Public Property SSPP_Min As Double
    Public Property PSP_Sum As Double 'Using only sum of Vitevitch's measures
    Public Property PSBP_Sum As Double
    Public Property S_PSP_Average As Double 'Using only average of Storkle's measures
    Public Property S_PSBP_Average As Double
    Public Property LanguageHomophones As List(Of String)
    Public Property LanguageHomophoneCount As Single
    Public Property FWPN_DensityProbability As Double 'Frequency Weighted Phonetic Neighbour Density Based Probability (FWPN_DensityProbability)
    Public Property PLD1Transcriptions As New List(Of String)
    Public Property FWON_DensityProbability As Double 'Frequency Weighted Orthographic Neighbour Density Based Probability (FWDN_DensityProbability)
    Public Property OLD1Spellings As New List(Of String)
    Public Property AllPossiblePoS As New List(Of Tuple(Of String, Double))
    Public Property AllOccurringLemmas As New List(Of Tuple(Of String, Double))
    Public Property NumberOfSenses As SByte
    Public Property Abbreviation As Boolean = False
    Public Property Acronym As Boolean = False
    Public Property AllPossibleSenses As New List(Of String)
    Public Property Sonographs_Letters As New List(Of String)
    Public Property Sonographs_Pronunciation As New List(Of String)

    Public ReadOnly Property SonographString As String
        Get

            Dim ItemCount As Integer = Math.Min(Sonographs_Letters.Count, Sonographs_Pronunciation.Count)

            Dim LocalSonographs As New List(Of String)
            For i = 0 To ItemCount - 1
                LocalSonographs.Add(Sonographs_Letters(i) & "-" & Sonographs_Pronunciation(i))
            Next

            If LocalSonographs.Count > 0 Then
                Return String.Join("|", LocalSonographs)
            Else
                Return ""
            End If

        End Get
    End Property

    Public Property GIL2P_OT As New List(Of Double)
    Public Property PIP2G_OT As New List(Of Double)
    Public Property G2P_OT As New List(Of Double)
    Public Property ForeignWord As Boolean = False
    'Public Property CorrectedSpelling As Boolean = False 'Hold a value indicating if the orthographic transcription has been corrected (manually or automatic)
    'Public Property CorrectedTranscription As Boolean = False 'Hold a value indicating if the phonetic transcription has been corrected (manually or automatic)
    Public Property ManuallyReveiwedCount As Integer = 0 'This summes all evaluation points upon export


    'Variables that are not read from the txt input file, but should be exported to the output file
    Public Property Tone As SByte
    Public Property MainStressSyllableIndex As SByte
    Public Property SecondaryStressSyllableIndex As SByte
    Public Property MostCommonPoS As Tuple(Of String, Double)
    Public Property MostCommonLemma As Tuple(Of String, Double)
    Public Property PLD1WordCount As Integer
    Public Property OLD1WordCount As Integer
    Public Property DiGraphCount As Integer
    Public Property TriGraphCount As Integer
    Public Property LongGraphemesCount As Integer

    'Isolation points
    Public Property OrthographicIsolationPoint As Integer = -1 '-1 Is set at the default value, indicating "not calculated"
    Public Property PhoneticIsolationPoint As Integer = -1 '-1 Is set at the default value, indicating "not calculated"


    'Support variables used to load / store / export variables rom output/input txt files

    Public Property PP As List(Of Double)
    Public Property PP_Phonemes As List(Of String)

    Public Property PSP As List(Of Double)
    Public Property PSP_Phonemes As List(Of String)

    Public Property PSBP As List(Of Double)
    Public Property PSBP_Phonemes As List(Of String)

    Public Property S_PSP As List(Of Double)
    Public Property S_PSP_Phonemes As List(Of String)

    Public Property S_PSBP As List(Of Double)
    Public Property S_PSBP_Phonemes As List(Of String)


    'Other support variables
    Public Property ManualEvaluations As New List(Of String)
    Public Property TranscriptionString As String
    Public Property Phonemes As New List(Of String)
    Public Property GroupHomophoneCount As Single
    Public Property GroupHomographCount As Single

    ''' <summary>
    ''' Can be used to temprarily store whether some particular or set of word list data has been loaded. Default value is True.
    ''' </summary>
    ''' <returns></returns>
    Public Property ContainsWordListData As Boolean = True 'This variable is used by the website to tell whether AFC-list data has been loaded for a particular word.

    'Word frequency
    Public ZipfValue_Word As Double


    ' Edit distance and phonologic distance
    ''' <summary>
    ''' Storing orthographic Levenshtein distance data. OLD data. Each Tuple containing: [OLDx-value, Spelling, ZipfValue]
    ''' </summary>
    Public OLDxData As New List(Of Tuple(Of Integer, String, Single)) '
    Public OLDx_Average As Double

    ''' <summary>
    ''' Storing phonetic Levenshtein distance data. OLD data. Each Tuple containing: [OLDx-value, PLD-Transcription string, ZipfValue]
    ''' </summary>
    Public PLDxData As New List(Of Tuple(Of Integer, String, Single)) '
    Public PLDx_Average As Double


    Public OLD_Data As OrthLD_Data_Type
    <Serializable>
    Public Class OrthLD_Data_Type
        Public Property OLD_Words As New List(Of WordGroup.LD_String)
        Public Property OLD20_Mean As Double
        Public Property OLD1_Count As Double
    End Class

    Public PLD_Data As PLD_Data_Type
    <Serializable>
    Public Class PLD_Data_Type
        Public Property PLD_Words As New List(Of WordGroup.LD_String)
        Public Property PLD20_Mean As Double
        Public Property PLD1_Count As Double
    End Class

    Public Property FrequencyWeightedDensity_OLD As Double

    Public Property CorrectedTranscription As Boolean = False

#End Region

#Region "SyllableClasses"

    ''' <summary>
    ''' Stores the syllables of a word. (N.B. Index of nucleus is a 1-based index)
    ''' </summary>
    <Serializable>
    Public Class ListOfSyllables
        Inherits List(Of Syllable)
        Public Property Tone As SByte '1 or 2
        Public Property MainStressSyllableIndex As SByte '1-based index
        Public Property SecondaryStressSyllableIndex As SByte '1-based index

        ''' <summary>
        ''' Reads the syllable class of a word and returns the next phoneme. If the current phoneme is the last in the word, and empty string will be returned.
        ''' </summary>
        ''' <param name="CurrentSyllableZeroBasedIndex"></param>
        ''' <param name="CurrentPhonemeIndex"></param>
        ''' <returns></returns>
        Public Function GetNextSound(ByVal CurrentSyllableZeroBasedIndex As Integer, ByVal CurrentPhonemeIndex As Integer) As String

            Try
                Dim OutputString = ""

                'If it's not the last sound in the current syllable, gets the next sound
                If CurrentPhonemeIndex < Me(CurrentSyllableZeroBasedIndex).Phonemes.Count - 1 Then
                    OutputString = Me(CurrentSyllableZeroBasedIndex).Phonemes(CurrentPhonemeIndex + 1)

                Else
                    'It is the last sound in the current syllable

                    'If it's not the last syllable, gets the first sound in the next syllable
                    If CurrentSyllableZeroBasedIndex < Me.Count - 1 Then
                        OutputString = Me(CurrentSyllableZeroBasedIndex + 1).Phonemes(0)
                    Else
                        'If it's the last sound in the last syllable, returns ""
                        OutputString = ""
                    End If
                End If

                Return OutputString

            Catch ex As Exception
                MsgBox(ex.ToString)
                Return Nothing
            End Try

        End Function

        ''' <summary>
        ''' Sets the value of the next sound in the syllables instance. If the current sound is the last sound in the last syllable, a new 
        ''' phoneme with the input value is added to the last syllable.
        ''' </summary>
        ''' <param name="CurrentSyllableZeroBasedIndex"></param>
        ''' <param name="CurrentPhonemeIndex"></param>
        ''' <param name="Value"></param>
        Public Sub SetNextSound(ByVal CurrentSyllableZeroBasedIndex As Integer, ByVal CurrentPhonemeIndex As Integer, ByRef Value As String)

            Try

                'If it's not the last sound in the current syllable, gets the next sound
                If CurrentPhonemeIndex < Me(CurrentSyllableZeroBasedIndex).Phonemes.Count - 1 Then
                    Me(CurrentSyllableZeroBasedIndex).Phonemes(CurrentPhonemeIndex + 1) = Value

                Else
                    'It is the last sound in the current syllable

                    'If it's not the last syllable, gets the first sound in the next syllable
                    If CurrentSyllableZeroBasedIndex < Me.Count - 1 Then
                        Me(CurrentSyllableZeroBasedIndex + 1).Phonemes(0) = Value
                    Else
                        'If it's the last sound in the last syllable, adds a new phoneme word finally
                        Me(CurrentSyllableZeroBasedIndex).Phonemes.Add(Value)
                    End If
                End If

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End Sub

        Public Function GetPreviousSound(ByVal CurrentSyllableZeroBasedIndex As Integer, ByVal CurrentPhonemeIndex As Integer) As String

            Try
                Dim OutputString = ""

                'If it's not the first sound in the current syllable, gets the previous sound
                If CurrentPhonemeIndex > 0 Then
                    OutputString = Me(CurrentSyllableZeroBasedIndex).Phonemes(CurrentPhonemeIndex - 1)

                Else
                    'It is the first sound in the current syllable

                    'If it's not the first syllable, gets the last sound in the previous syllable
                    If CurrentSyllableZeroBasedIndex > 0 Then
                        OutputString = Me(CurrentSyllableZeroBasedIndex - 1).Phonemes(Me(CurrentSyllableZeroBasedIndex - 1).Phonemes.Count - 1)
                    Else
                        'If it's the last sound in the last syllable, returns ""
                        OutputString = ""
                    End If
                End If

                Return OutputString

            Catch ex As Exception
                MsgBox(ex.ToString)
                Return Nothing
            End Try

        End Function

        ''' <summary>
        ''' Sets the value of the previous sound in the syllables instance. If the current sound is the first sound in the first syllable, a new 
        ''' phoneme with the input value is inserted word initially.
        ''' </summary>
        ''' <param name="CurrentSyllableZeroBasedIndex"></param>
        ''' <param name="CurrentPhonemeIndex"></param>
        ''' <param name="Value"></param>
        Public Sub SetPreviousSound(ByVal CurrentSyllableZeroBasedIndex As Integer, ByVal CurrentPhonemeIndex As Integer, ByRef Value As String)

            Try

                'If it's not the first sound in the current syllable, gets the next sound
                If CurrentPhonemeIndex > 0 Then
                    Me(CurrentSyllableZeroBasedIndex).Phonemes(CurrentPhonemeIndex - 1) = Value

                Else
                    'It is the first sound in the current syllable

                    'If it's not the first syllable, sets the last sound in the previous syllable
                    If CurrentSyllableZeroBasedIndex > 0 Then
                        Me(CurrentSyllableZeroBasedIndex - 1).Phonemes(Me(CurrentSyllableZeroBasedIndex - 1).Phonemes.Count - 1) = Value
                    Else
                        'If it's the first sound in the first syllable, inserts a word new phoneme word initially 
                        Me(0).Phonemes.Insert(0, Value)
                    End If
                End If

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End Sub

        ''' <summary>
        ''' Translates a whole-word phoneme index of a word to the syllable and syllable phoneme indices.
        ''' </summary>
        ''' <param name="WordPhonemeIndex"></param>
        ''' <param name="SyllableIndex"></param>
        ''' <param name="SyllablePhonemeIndex"></param>
        Public Sub Translate_WordPhonemeIndex_To_SyllableIndices(ByVal WordPhonemeIndex As Integer, ByRef SyllableIndex As Integer, ByRef SyllablePhonemeIndex As Integer)


            'Determines which syllable it is
            Dim PhonemeCount As Integer = 0
            Dim LastSyllableCount As Integer = 0
            For syll = 0 To Me.Count - 1

                If WordPhonemeIndex >= PhonemeCount Then
                    SyllableIndex = syll
                Else
                    Exit For
                End If

                'Accumulating phoneme count, starting from the beginning of the word, until the correct syllable index has been found
                PhonemeCount += Me(syll).Phonemes.Count

                LastSyllableCount = Me(syll).Phonemes.Count

            Next

            'Detemines which phoneme it is
            SyllablePhonemeIndex = WordPhonemeIndex - (PhonemeCount - LastSyllableCount)

        End Sub

        ''' <summary>
        ''' Translates a syllable and syllable phoneme index to a whole-word (only) phoneme array index of a word.
        ''' </summary>
        ''' <param name="SyllableIndex"></param>
        ''' <param name="SyllablePhonemeIndex"></param>
        Public Function Translate_SyllableIndices_To_WordPhonemeIndex(ByRef SyllableIndex As Integer, ByRef SyllablePhonemeIndex As Integer)

            Dim Result As Integer = 0

            'Counting the phonemes in the preceding syllables
            For syll = 0 To SyllableIndex - 1
                Result += Me(syll).Phonemes.Count
            Next

            'Adding the preceding phonemes in the current syllable, and substracting 1, since the phoneme index is zero-based
            Result += SyllablePhonemeIndex - 1

            Return Result

        End Function

        ''' <summary>
        ''' Returns the total number of phonemes in all syllables of the word
        ''' </summary>
        ''' <param name="IncludeZeroPhoneme">If set to true, any zero phonemes are not counted.</param>
        ''' <returns></returns>
        Public Function GetWordPhonemeCount(Optional ByRef IncludeZeroPhoneme As Boolean = True) As Integer

            Dim PhonemeCount As Integer = 0
            If IncludeZeroPhoneme = True Then

                For syll = 0 To Me.Count - 1
                    PhonemeCount += Me(syll).Phonemes.Count
                Next

            Else
                For syll = 0 To Me.Count - 1
                    For p = 0 To Me(syll).Phonemes.Count - 1
                        If Not Me(syll).Phonemes(p) = ZeroPhoneme Then PhonemeCount += 1
                    Next
                Next
            End If

            Return PhonemeCount

        End Function

        ''' <summary>
        ''' Returns a new ListOfSyllables which is a deep copy of the original.
        ''' </summary>
        ''' <returns>Returns a new ListOfSyllables which is a deep copy of the original.</returns>
        Public Function CreateCopy() As ListOfSyllables

            'Serializing to memorystream
            Dim serializedMe As New MemoryStream
            Dim serializer As New BinaryFormatter
            serializer.Serialize(serializedMe, Me)

            'Deserializing Me to a new object
            serializedMe.Position = 0
            Dim newListOfSyllables As ListOfSyllables = CType(serializer.Deserialize(serializedMe), ListOfSyllables)
            serializedMe.Close()

            'Returning the new object
            Return newListOfSyllables

        End Function

    End Class



    ''' <summary>
    ''' Class that holds syllable information. The index indexOfNuclues is 1-based.
    ''' </summary>
    <Serializable>
    Public Class Syllable
        Public Property IsStressed As Boolean = False
        Public Property CarriesSecondaryStress As Boolean = False
        Public Property Phonemes As New List(Of String)
        ''' <summary>
        ''' Returns phonemes which have reduced vowel quality, according to the VowelAllophoneReductionList dictionary.
        ''' </summary>
        ''' <param name="PhonemeIndex"></param>
        ''' <param name="RemovePhoneticLengthInSecondarilyStressedSyllables">If set to true, any phonetic length character attached to the current phoneme will be removed. If the phoneme is a vowel, length removal will take place prior to vowel neutralization.</param>
        ''' <returns></returns>
        Public ReadOnly Property SurfacePhones(PhonemeIndex As Integer,
                                                   Optional RemovePhoneticLengthInSecondarilyStressedSyllables As Boolean = False) As String 'Optional RemoveLengthCharacterAfterVowelReduction As Boolean = False) As String
            Get
                Dim ReturnString As String = Phonemes(PhonemeIndex)

                If CarriesSecondaryStress = True And RemovePhoneticLengthInSecondarilyStressedSyllables = True Then
                    ReturnString = ReturnString.Replace(PhoneticLength, "")
                End If

                'Replacing any length reduced long vowels (except /a/) with their short vowel equivalent
                If VowelAllophoneReductionList.ContainsKey(ReturnString) Then
                    ReturnString = VowelAllophoneReductionList(ReturnString)
                End If

                'If RemoveLengthCharacterAfterVowelReduction = True Then
                'ReturnString = ReturnString.Replace(PhoneticLength, "")
                'End If

                Return ReturnString
            End Get
        End Property
        Public Property SyllableLength As SByte
        Public Property IndexOfNuclues As SByte = 0
        Public Property LengthOfOnset As SByte
        Public Property LengthOfCoda As SByte
        Public Property StringRepressentation As String
        Public Property AmbigousOnset As Boolean = False
        Public Property AmbigousCoda As Boolean = False
        Public Property SyllableCodaType As SyllableTypes

        Public Function CreateCopy() As Syllable

            'Serializing to memorystream
            Dim serializedMe As New MemoryStream
            Dim serializer As New BinaryFormatter
            serializer.Serialize(serializedMe, Me)

            'Deserializing Me to a new object
            serializedMe.Position = 0
            Dim SyllableCopy As Syllable = CType(serializer.Deserialize(serializedMe), Syllable)
            serializedMe.Close()

            'Returning the new object
            Return SyllableCopy

        End Function

    End Class

    ''' <summary>
    ''' A descriptive enumerator of syllable openness. N.B. this does not tell whether the syllable coda has ambigous consonants or not.
    ''' </summary>
    Public Enum SyllableTypes
        Closed
        Open
    End Enum
#End Region

#Region "Phonology"


    ''' <summary>
    ''' Adds zero phones before word initial, and after word final, vowels
    ''' </summary>
    Public Sub CheckAndAddZeroPhones(Optional ByRef ErrorsCorrected As Integer = 0, Optional ByRef NonCorrectedErrors As Integer = 0, Optional ByVal LogChanges As Boolean = True)



        If Syllables.Count > 0 Then

            'Checks initial 0+non vowel
            If Syllables(0).Phonemes(0) = ZeroPhoneme Then
                Try
                    If Not SwedishVowels_IPA.Contains(Syllables(0).Phonemes(1)) Then
                        ErrorsCorrected += 1
                        If LogChanges = True Then SendInfoToLog("Pre-addition: " & String.Join(" ", BuildExtendedIpaArray) & vbTab & OrthographicForm, "Init_ZeroP+Cons")

                        Syllables(0).Phonemes.RemoveAt(0)
                        DetermineSyllableIndices()
                        If LogChanges = True Then SendInfoToLog(String.Join(" ", BuildExtendedIpaArray) & vbTab & OrthographicForm, "Init_ZeroP+Cons")
                    End If
                Catch ex As Exception
                    NonCorrectedErrors += 1
                    If LogChanges = True Then SendInfoToLog(String.Join(" ", BuildExtendedIpaArray) & vbTab & OrthographicForm, "UnableToAddZeroPhoneme")
                End Try
            End If

            'Checks initial vowel without preceding 0
            If SwedishVowels_IPA.Contains(Syllables(0).Phonemes(0)) Then
                ErrorsCorrected += 1
                If LogChanges = True Then SendInfoToLog("Pre-addition: " & String.Join(" ", BuildExtendedIpaArray) & vbTab & OrthographicForm, "Init_Vowel_NoZeroP")
                Syllables(0).Phonemes.Insert(0, ZeroPhoneme)
                DetermineSyllableIndices()
                If LogChanges = True Then SendInfoToLog(String.Join(" ", BuildExtendedIpaArray) & vbTab & OrthographicForm, "Init_Vowel_NoZeroP")
            End If

            'Checks final non vowel + 0
            If Syllables(Syllables.Count - 1).Phonemes(Syllables(Syllables.Count - 1).Phonemes.Count - 1) = ZeroPhoneme Then
                Try

                    If Not SwedishVowels_IPA.Contains(Syllables(Syllables.Count - 1).Phonemes(Syllables(Syllables.Count - 1).Phonemes.Count - 2)) Then
                        ErrorsCorrected += 1
                        If LogChanges = True Then SendInfoToLog("Pre-addition: " & String.Join(" ", BuildExtendedIpaArray) & vbTab & OrthographicForm, "Fin_Cons+ZeroP")
                        Syllables(Syllables.Count - 1).Phonemes.RemoveAt(Syllables(Syllables.Count - 1).Phonemes.Count - 1)
                        DetermineSyllableIndices()
                        If LogChanges = True Then SendInfoToLog(String.Join(" ", BuildExtendedIpaArray) & vbTab & OrthographicForm, "Fin_Cons+ZeroP")
                    End If
                Catch ex As Exception
                    NonCorrectedErrors += 1
                    If LogChanges = True Then SendInfoToLog(String.Join(" ", BuildExtendedIpaArray) & vbTab & OrthographicForm, "UnableToAddZeroPhoneme")
                End Try
            End If

            'Checks final vowel without following 0
            If SwedishVowels_IPA.Contains(Syllables(Syllables.Count - 1).Phonemes(Syllables(Syllables.Count - 1).Phonemes.Count - 1)) Then
                ErrorsCorrected += 1
                If LogChanges = True Then SendInfoToLog("Pre-addition: " & String.Join(" ", BuildExtendedIpaArray) & vbTab & OrthographicForm, "Fin_Vowel_NoZeroP")
                Syllables(Syllables.Count - 1).Phonemes.Add(ZeroPhoneme)
                DetermineSyllableIndices()
                If LogChanges = True Then SendInfoToLog(String.Join(" ", BuildExtendedIpaArray) & vbTab & OrthographicForm, "Fin_Vowel_NoZeroP")
            End If
        End If

    End Sub


    ''' <summary>
    ''' Removes all zero phonemes stored in the indicated syllables class
    ''' </summary>
    ''' <param name="AlternativeSyllabification">If set to true the removal of zero phonemes will be performed in the alternative syllabification syllables, if set to false it will be performed on the Word.Syllables instance:</param>
    Public Sub RemoveZeroPhonemes(ByVal AlternativeSyllabification As Boolean)

        If AlternativeSyllabification = False Then

            For Each CurrentSyllable In Syllables
                Dim Ph As Integer = 0
                Do Until Ph > CurrentSyllable.Phonemes.Count - 1
                    If CurrentSyllable.Phonemes(Ph) = ZeroPhoneme Then
                        CurrentSyllable.Phonemes.RemoveAt(Ph)
                    Else
                        Ph += 1
                    End If
                Loop
            Next

        Else

            For Each CurrentSyllable In Syllables_AlternateSyllabification
                Dim Ph As Integer = 0
                Do Until Ph > CurrentSyllable.Phonemes.Count - 1
                    If CurrentSyllable.Phonemes(Ph) = ZeroPhoneme Then
                        CurrentSyllable.Phonemes.RemoveAt(Ph)
                    Else
                        Ph += 1
                    End If
                Loop
            Next

        End If

        DetermineSyllableIndices(AlternativeSyllabification)

    End Sub


    ''' <summary>
    ''' Removes the first of two adjacent phonetic characters in the syllables array. Returns the number of removed phonemes.
    ''' </summary>
    ''' <param name="RemoveConsonants"></param>
    ''' <param name="RemoveVowels"></param>
    ''' <param name="LogWords"></param>
    ''' <param name="DisregardPhoneticLengthCharacter">If set to true, length is neutralised when comparing the phonemes. The first phoneme is still removed, but if any of the two phonemes have phonetic length, it will be retained in the remaining phoneme.</param>
    ''' <returns></returns>
    Public Function RemoveDoublePhoneticCharacters(Optional ByVal RemoveConsonants As Boolean = True, Optional ByVal RemoveVowels As Boolean = False,
                                                  Optional LogWords As Boolean = False, Optional ByVal DisregardPhoneticLengthCharacter As Boolean = False,
                                                   Optional IncludeSAMPA As Boolean = True) As Integer

        Dim RemovedPhonemes As Integer = 0


        For syllable = 0 To Syllables.Count - 1

            Dim PhonemesToRemove As New List(Of RemovalIndexAndLength)

            For phoneme = 0 To Syllables(syllable).Phonemes.Count - 1

                Dim DoubleDetected As Integer = 0

                'If it's not the last phoneme in the syllable
                If Not phoneme = Syllables(syllable).Phonemes.Count - 1 Then

                    'Testing if the phoneme is repeated
                    If DisregardPhoneticLengthCharacter = False Then
                        If Syllables(syllable).Phonemes(phoneme) = Syllables(syllable).Phonemes(phoneme + 1) Then
                            DoubleDetected += 1
                        End If
                    Else
                        If Syllables(syllable).Phonemes(phoneme).Replace(PhoneticLength, "") = Syllables(syllable).Phonemes(phoneme + 1).Replace(PhoneticLength, "") Then
                            DoubleDetected += 1
                        End If
                    End If


                Else
                    'It is the last phoneme in the current syllable
                    'If it's not the last syllable, looking at the first phoneme in the next syllable
                    If Not syllable = Syllables.Count - 1 Then

                        'Testing if the phoneme is repeated
                        If DisregardPhoneticLengthCharacter = False Then
                            If Syllables(syllable).Phonemes(phoneme) = Syllables(syllable + 1).Phonemes(0) Then
                                DoubleDetected += 1
                            End If
                        Else
                            If Syllables(syllable).Phonemes(phoneme).Replace(PhoneticLength, "") = Syllables(syllable + 1).Phonemes(0).Replace(PhoneticLength, "") Then
                                DoubleDetected += 1
                            End If
                        End If
                    End If
                End If


                'Mark index for removal
                If DoubleDetected > 0 Then
                    'If it is a consonant
                    If SwedishConsonants_IPA.Contains(Syllables(syllable).Phonemes(phoneme)) Then
                        If RemoveConsonants = True Then
                            'Removes the leftmost phoneme 

                            Dim newRemoval As New RemovalIndexAndLength
                            newRemoval.RemovalIndex = phoneme
                            If Syllables(syllable).Phonemes(phoneme).Contains(PhoneticLength) Then newRemoval.HasLength = True
                            PhonemesToRemove.Add(newRemoval)

                        End If

                        If LogWords = True Then
                            If IncludeSAMPA = True Then
                                SendInfoToLog(OrthographicForm & vbTab & String.Concat(BuildExtendedIpaArray()) & vbTab &
                                          "Double sound:  " & vbTab & Syllables(syllable).Phonemes(phoneme), "DoubleConsonants")
                            Else
                                SendInfoToLog(OrthographicForm & vbTab & String.Concat(BuildExtendedIpaArray()) & vbTab &
                                          "Double sound:  " & vbTab & Syllables(syllable).Phonemes(phoneme), "DoubleConsonants")
                            End If
                        End If

                    End If

                    'If it's a vowel
                    If SwedishVowels_IPA.Contains(Syllables(syllable).Phonemes(phoneme)) Then
                        If RemoveVowels = True Then

                            'Removes the leftmost phoneme 
                            Dim newRemoval As New RemovalIndexAndLength
                            newRemoval.RemovalIndex = phoneme
                            If Syllables(syllable).Phonemes(phoneme).Contains(PhoneticLength) Then newRemoval.HasLength = True
                            PhonemesToRemove.Add(newRemoval)

                        End If

                        If LogWords = True Then
                            If IncludeSAMPA = True Then
                                SendInfoToLog(OrthographicForm & vbTab & String.Concat(BuildExtendedIpaArray()) & vbTab &
                                                  "Double sound: " & vbTab & Syllables(syllable).Phonemes(phoneme), "DoubleVowels")

                            Else
                                SendInfoToLog(OrthographicForm & vbTab & String.Concat(BuildExtendedIpaArray()) & vbTab &
                                                  "Double sound: " & vbTab & Syllables(syllable).Phonemes(phoneme), "DoubleVowels")

                            End If

                        End If

                    End If
                End If
            Next

            'Removing phonemes, in reversed order
            PhonemesToRemove.Sort()
            PhonemesToRemove.Reverse()
            Dim NextSyllableIsChanged As Boolean = False
            For phoneme = 0 To PhonemesToRemove.Count - 1

                'Adding length marker to the remaining phoneme (the rightmost phoneme/the next phoneme after the one to be removed), 
                'if the deleted segment had length, DisregardPhoneticLengthCharacter is true, 
                'and the remaining segment is not already long
                If DisregardPhoneticLengthCharacter = True And PhonemesToRemove(phoneme).HasLength = True Then
                    Dim RightMostSegment As String = Syllables.GetNextSound(syllable, PhonemesToRemove(phoneme).RemovalIndex)
                    If RightMostSegment <> "" Then
                        If Not RightMostSegment.Contains(PhoneticLength) Then
                            'Adding phonetic length
                            Syllables.SetNextSound(syllable, PhonemesToRemove(phoneme).RemovalIndex, RightMostSegment & PhoneticLength)
                        End If
                    End If
                End If

                'Removes the phoneme at the specified index if this is within the range of the phoneme array,
                'If it is outside that range it will indicat that it is the first phoneme in the next syllable that should be removed
                If PhonemesToRemove(phoneme).RemovalIndex <= Syllables(syllable).Phonemes.Count - 1 Then
                    Syllables(syllable).Phonemes.RemoveAt(PhonemesToRemove(phoneme).RemovalIndex)
                Else
                    Syllables(syllable + 1).Phonemes.RemoveAt(0)
                    NextSyllableIsChanged = True
                End If

            Next

            If PhonemesToRemove.Count > 0 Then
                'Noting that the word is corrected
                CorrectedTranscription = True

                'Updating syllable internal structure of the current syllable, and if needed also the following
                DetermineSyllableIndices(syllable)
                If NextSyllableIsChanged = True Then DetermineSyllableIndices(syllable + 1)
            End If

            'Accumulating removed phoneme count
            RemovedPhonemes += PhonemesToRemove.Count

        Next

        Return RemovedPhonemes

    End Function

    Private Class RemovalIndexAndLength
        Implements IComparable

        Property RemovalIndex As Integer
        Property HasLength As Boolean = False

        Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo

            If Not TypeOf (obj) Is RemovalIndexAndLength Then
                Throw New ArgumentException()
            Else

                Dim ComparisonValue As RemovalIndexAndLength = DirectCast(obj, RemovalIndexAndLength)

                If Me.RemovalIndex < ComparisonValue.RemovalIndex Then
                    Return -1
                ElseIf Me.RemovalIndex = ComparisonValue.RemovalIndex Then
                    Return 0
                Else
                    Return 1
                End If
            End If
        End Function
    End Class

    ''' <summary>
    ''' Determines if the words in the word group are of the type CVC, CV, CVCC, CVCVC, etc. The result is stored in PhonotacticType, as well as returned from the function. 
    ''' </summary>
    Public Function SetWordPhonotacticType(Optional ByRef PhonemesNeitherConsonantsOrVowels As Integer = 0,
        Optional ByRef PhonemesNeitherConsonantsOrVowelTypes As SortedList(Of String, Integer) = Nothing) As String

        PhonotacticType = ""
        Dim ReducedPhoneticForm = BuildExtendedIpaArray(, True,,,,, False)

        For ph = 0 To ReducedPhoneticForm.Count - 1
            If SwedishConsonants_IPA.Contains(ReducedPhoneticForm(ph)) Then
                PhonotacticType &= "C"
            ElseIf SwedishVowels_IPA.Contains(ReducedPhoneticForm(ph)) Then
                PhonotacticType &= "V"
            Else
                PhonemesNeitherConsonantsOrVowels += 1
                If PhonemesNeitherConsonantsOrVowelTypes Is Nothing Then PhonemesNeitherConsonantsOrVowelTypes = New SortedList(Of String, Integer)
                If Not PhonemesNeitherConsonantsOrVowelTypes.ContainsKey(ReducedPhoneticForm(ph)) Then
                    PhonemesNeitherConsonantsOrVowelTypes.Add(ReducedPhoneticForm(ph), 1)
                Else
                    PhonemesNeitherConsonantsOrVowelTypes(ReducedPhoneticForm(ph)) += 1
                End If
            End If
        Next

        Return PhonotacticType

    End Function



    ''' <summary>
    ''' Creates IPA, extended IPA and Phonemes Array for the words, using the Syllables property. Also counting phonemes
    ''' </summary>
    Public Sub GeneratePhoneticForms()

        'Recreating the phonetic transcription
        Dim ExtendedIpaArray = BuildExtendedIpaArray()

        'Creating a word Extended IPA form
        TranscriptionString = String.Concat(ExtendedIpaArray)
        TranscriptionString = TranscriptionString.Replace(IpaSyllableBoundary, "")


    End Sub

    Public Function CountPhonemes(Optional ByRef IncludeZeroPhoneme As Boolean = False, Optional ByVal TemporarilyAddZeroPhonemes As Boolean = False)

        If TemporarilyAddZeroPhonemes = False Then
            'Creating a phonetic form array, with only phoneme characters
            Dim CurrentPhonemes As List(Of String) = BuildExtendedIpaArray(, True,,,,, IncludeZeroPhoneme)

            'Counting phonemes
            Return CurrentPhonemes.Count
        Else

            'Creating a new word in which the zero phonemes are added
            Dim TempWord As Word = Me.CreateCopy

            'Adding Zero phones to the temporary word
            TempWord.CheckAndAddZeroPhones(,, False)

            Return TempWord.CountPhonemes(IncludeZeroPhoneme, False)

        End If

    End Function

    ''' <summary>
    ''' This function creates an array containing all available phonetic characters in the syllables of a word.
    ''' </summary>
    ''' <param name="IncludeAmbiguityMarkers">Special markers inserted into the first phoneme string in each syllable, to incicate ambigous syllable bondaries.</param>
    ''' <param name="IncludeOnlyPhonemes">If set to true, only the phonemes stored in the Syllables of the words will be returned.</param>
    ''' <param name="ReduceSecondaryStress">If set to true, the stress marker and phonetic length characters will be removed from syllables with secondary stress. This parameter has no effect if IncludeOnlyPhonemes is set to true.</param>''' 
    ''' <param name="ReturnSurfacePhones"></param>
    ''' <param name="UseAlternateSyllabification">Reads the Word.Syllables_AlternateSyllabification instead of Word.Syllables. If Word.Syllables_AlternateSyllabification is nothing, an empty list of string is returned.</param>
    ''' <param name="IncludeZeroPhonemes">If set to false, any Zero phonemes in the transcription will be removed. (However, if no zero phonemes exist in the syllable transcriptions, they are not added, even if IncludeZeroPhonemes is set to true.)</param>
    ''' <returns></returns>
    Public Function BuildExtendedIpaArray(Optional ByVal IncludeAmbiguityMarkers As Boolean = False,
                                              Optional ByVal IncludeOnlyPhonemes As Boolean = False,
                                              Optional ReduceSecondaryStress As Boolean = False,
                                              Optional ReturnSurfacePhones As Boolean = False,
                                              Optional ByRef UseAlternateSyllabification As Boolean = False,
                                              Optional ByVal IncludeSyllableBoundaries As Boolean = True,
                                              Optional ByVal IncludeZeroPhonemes As Boolean = True) As List(Of String)


        Dim SyllablesToUse As ListOfSyllables
        If UseAlternateSyllabification = False Then
            SyllablesToUse = Syllables
        Else
            If Syllables_AlternateSyllabification Is Nothing Then

                'Returns an empty list of string, if there exists no Syllables_AlternateSyllabification.
                Return New List(Of String)

                'Code previous to 2017-11-07. This code read Syllables instead of Syllables_AlternateSyllabification if Syllables_AlternateSyllabification was Nothing. And alterred UseAlternateSyllabification to False upon return if Syllables_AlternateSyllabification was Nothing. The calling code then had to check for this.
                'SyllablesToUse = Syllables
                'UseAlternateSyllabification = False
            Else
                SyllablesToUse = Syllables_AlternateSyllabification
            End If
        End If


        Dim IPAform As New List(Of String)

        If SyllablesToUse.Count > 0 Then

            For syllable = 0 To SyllablesToUse.Count - 1

                Dim CurrentSyllableIPAform As New List(Of String)

                'Adding phonemes to the output array
                If ReturnSurfacePhones = False Then
                    For phoneme = 0 To SyllablesToUse(syllable).Phonemes.Count - 1
                        CurrentSyllableIPAform.Add(SyllablesToUse(syllable).Phonemes(phoneme))
                    Next
                Else

                    For phoneme = 0 To SyllablesToUse(syllable).Phonemes.Count - 1
                        CurrentSyllableIPAform.Add(SyllablesToUse(syllable).SurfacePhones(phoneme, ReduceSecondaryStress))
                    Next
                End If

                If IncludeOnlyPhonemes = False Then

                    'Inserting stress / accent markers to stressed syllables, in the position just before the nucleus,
                    'or before the first phoneme if en error in nucleus index exists
                    If SyllablesToUse(syllable).IsStressed = True Then
                        If syllable = MainStressSyllableIndex - 1 Then
                            Select Case Tone
                                Case 1
                                    If SyllablesToUse(syllable).IndexOfNuclues > 0 Then
                                        CurrentSyllableIPAform.Insert(SyllablesToUse(syllable).IndexOfNuclues - 1, IpaMainStress)
                                    Else
                                        CurrentSyllableIPAform.Insert(0, IpaMainStress)
                                    End If

                                Case 2
                                    If SyllablesToUse(syllable).IndexOfNuclues > 0 Then
                                        CurrentSyllableIPAform.Insert(SyllablesToUse(syllable).IndexOfNuclues - 1, IpaMainSwedishAccent2)
                                    Else
                                        CurrentSyllableIPAform.Insert(0, IpaMainSwedishAccent2)
                                    End If

                                Case Else
                                    MsgBox("Error")
                            End Select
                        End If

                        'Adding secondary stress markers 
                        If syllable = SecondaryStressSyllableIndex - 1 Then
                            If SyllablesToUse(syllable).IndexOfNuclues > 0 Then
                                CurrentSyllableIPAform.Insert(SyllablesToUse(syllable).IndexOfNuclues - 1, IpaSecondaryStress)
                            Else
                                CurrentSyllableIPAform.Insert(0, IpaSecondaryStress)
                            End If
                        End If

                    End If
                End If

                'Adding ambiguity markers in the first phoneme string
                If IncludeAmbiguityMarkers = True Then
                    If SyllablesToUse(syllable).AmbigousOnset = True Then
                        SyllablesToUse(syllable).Phonemes(0) &= AmbiguosOnsetMarker
                    End If
                    If SyllablesToUse(syllable).AmbigousCoda = True Then
                        SyllablesToUse(syllable).Phonemes(0) &= AmbiguosCodaMarker
                    End If
                End If

                'Adding the current syllable IPA form, and a syllable boundary marker
                For CurrSyllableItem = 0 To CurrentSyllableIPAform.Count - 1
                    IPAform.Add(CurrentSyllableIPAform(CurrSyllableItem))
                Next
                If IncludeOnlyPhonemes = False Then
                    If IncludeSyllableBoundaries = True Then
                        IPAform.Add(IpaSyllableBoundary)
                    End If
                End If
            Next

            'Removes the last syllable boundary
            Try
                If IncludeOnlyPhonemes = False Then
                    If IncludeSyllableBoundaries = True Then
                        IPAform.RemoveAt(IPAform.Count - 1)
                    End If
                End If

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If

        'Removing the first secondary stress marker, in the word (as only the first secondary stressed syllable has lost its length marking above)
        If ReduceSecondaryStress = True Then
            If SecondaryStressSyllableIndex <> 0 Then
                IPAform.Remove(IpaSecondaryStress)
            End If
        End If

        If IncludeZeroPhonemes = False Then
            'Removes zero phonemes
            Dim PhIndex As Integer = 0
            Do Until PhIndex > IPAform.Count - 1
                If IPAform(PhIndex) = ZeroPhoneme Then
                    IPAform.RemoveAt(PhIndex)
                Else
                    PhIndex += 1
                End If
            Loop
        End If

        Return IPAform

    End Function

    ''' <summary>
    ''' Builds a standard IPA transcription used for Calculations of PLD1 words, Homophone and homograph detection.
    ''' </summary>
    ''' <param name="UseAlternateSyllabification"></param>
    ''' <returns></returns>
    Public Function BuildReducedIpaArray(Optional ByRef UseAlternateSyllabification As Boolean = False,
                                             Optional ByVal IncludeZeroPhonemes As Boolean = False) As List(Of String)

        Dim Output As List(Of String) = BuildExtendedIpaArray(False, False, True, True, UseAlternateSyllabification, False)

        If IncludeZeroPhonemes = False Then
            'Removes zero phonemes
            Dim PhIndex As Integer = 0
            Do Until PhIndex > Output.Count - 1
                If Output(PhIndex) = ZeroPhoneme Then
                    Output.RemoveAt(PhIndex)
                Else
                    PhIndex += 1
                End If
            Loop
        End If

        Return Output

    End Function

    ''' <summary>
    ''' Builds a transcription string for comparisons of PLD1 (phonetic neighborhood)
    ''' </summary>
    ''' <returns></returns>
    Public Function BuildPLD1TypeTranscription() As List(Of String)

        Dim CurrentWordPhonemes As List(Of String) = BuildReducedIpaArray(False, False)

        'Removing stress markers
        Dim PhIndex As Integer = 0
        Do Until PhIndex > CurrentWordPhonemes.Count - 1
            If SwedishStressList.Contains(CurrentWordPhonemes(PhIndex)) Then
                CurrentWordPhonemes.RemoveAt(PhIndex)
            Else
                PhIndex += 1
            End If
        Loop

        'Puts main stress info initially (secondary stress is ignored)
        'Putting/inserting stressmarkers to denote tone on index 0, and the mainstress syllable index on index 1
        CurrentWordPhonemes.Insert(0, MainStressSyllableIndex)

        If Tone = 1 Then
            CurrentWordPhonemes.Insert(0, IpaMainStress)
        ElseIf Tone = 2 Then
            CurrentWordPhonemes.Insert(0, IpaMainSwedishAccent2)
        End If

        Return CurrentWordPhonemes

    End Function



    ''' <summary>
    ''' Marks syllable structure errors, by looking at the syllables array. 
    ''' Identifying errors types: WordWithoutSyllable, SyllableWithoutNucleus, NonConsonantInOnset, NonVowelInNucleus, NonConsonantInCoda
    ''' </summary>
    Public Function MarkSyllableStructureErrors() As Integer

        Dim TotalErrorsFound As Integer = 0

        'WordWithoutSyllable
        If Syllables.Count = 0 Then
            ManualEvaluations.Add(PhoneticMarkingsTypes.WordWithoutSyllable.ToString)
            CorrectedTranscription = True
            TotalErrorsFound += 1
        End If

        'Word without primary stress
        If MainStressSyllableIndex = 0 Or Tone = 0 Then
            ManualEvaluations.Add(PhoneticMarkingsTypes.NoMainStress.ToString)
            CorrectedTranscription = True
            TotalErrorsFound += 1
        End If

        'Testing for conflicting stress positions
        If MainStressSyllableIndex = SecondaryStressSyllableIndex And MainStressSyllableIndex <> 0 Then
            ManualEvaluations.Add(PhoneticMarkingsTypes.StressPositionConflict.ToString)
            CorrectedTranscription = True
            TotalErrorsFound += 1
        End If

        For syllable = 0 To Syllables.Count - 1

            If Syllables(syllable).IndexOfNuclues = 0 Then

                'SyllableWithoutNucleus
                ManualEvaluations.Add(PhoneticMarkingsTypes.SyllableWithoutNuclues.ToString & ", index: " & syllable)
                CorrectedTranscription = True
                TotalErrorsFound += 1

            Else
                'Doing the rest of the checks only if there is an established nucleus

                'Looing for syllables with incorrect type of sound in the different parts of the syllable,
                'Consonants or zero-phoneme are allowed in onset and coda, and only vowels are allowed in the nucleus

                'Testing the onset
                For phoneme = 0 To Syllables(syllable).IndexOfNuclues - 2
                    If Not SwedishConsonants_IPA.Contains(Syllables(syllable).Phonemes(phoneme)) Then
                        If Not Syllables(syllable).Phonemes(phoneme) = ZeroPhoneme Then
                            ManualEvaluations.Add(PhoneticMarkingsTypes.NonConsonantInOnset.ToString & ", index: " & syllable)
                            CorrectedTranscription = True
                            TotalErrorsFound += 1
                        End If
                    End If
                Next

                'Testing the nucleus
                If Not SwedishVowels_IPA.Contains(Syllables(syllable).Phonemes(Syllables(syllable).IndexOfNuclues - 1)) Then
                    ManualEvaluations.Add(PhoneticMarkingsTypes.NonVowelInNucleus.ToString & ", index: " & syllable)
                    CorrectedTranscription = True
                    TotalErrorsFound += 1
                End If

                'Testing the coda
                For phoneme = Syllables(syllable).IndexOfNuclues To Syllables(syllable).Phonemes.Count - 1
                    If Not Syllables(syllable).Phonemes(phoneme) = ZeroPhoneme Then
                        If Not SwedishConsonants_IPA.Contains(Syllables(syllable).Phonemes(phoneme)) Then
                            ManualEvaluations.Add(PhoneticMarkingsTypes.NonConsonantInCoda.ToString & ", index: " & syllable)
                            CorrectedTranscription = True
                            TotalErrorsFound += 1
                        End If
                    End If
                Next
            End If
        Next

        Return TotalErrorsFound

    End Function



    ''' <summary>
    ''' Marks phonetic form errors, by looking at the syllables array. 
    ''' Identified errors types: WordWithoutSyllable, WordWithoutMainStress, SyllableWithoutNucleus, SyllableWeightError
    ''' This sub requires DetermineSyllableOpenness to be allready run, and that long consonants are marked phonetically.
    ''' </summary>
    Public Function MarkSyllableWeightErrors(Optional ByRef LogAllErrors As Boolean = False)

        Dim TotalErrorsFound As Integer = 0

        For syllable = 0 To Syllables.Count - 1

            If Syllables(syllable).IndexOfNuclues = 0 Then

                ManualEvaluations.Add(PhoneticMarkingsTypes.SyllableWithoutNuclues.ToString & ", index: " & syllable)
                CorrectedTranscription = True
                TotalErrorsFound += 1

                If LogAllErrors = True Then SendInfoToLog(OrthographicForm & vbTab & String.Join(" ", BuildExtendedIpaArray), "SyllableWeightErrorWords_NoNucleus")

            Else

                If Syllables(syllable).IsStressed = True Then

                    'Doing the rest of the checks only if there is an established nucleus

                    'Looking for stressed syllables with incorrect weight (Cf. Riad p 160)
                    'A short stressed nucleus should be followed by a long consonant, and a long stressed nucleus should not be followed by a long consonant

                    'Getting the nucleus
                    Dim Nucleus As String = Syllables(syllable).Phonemes(Syllables(syllable).IndexOfNuclues - 1)
                    Dim NextSound As String = Syllables.GetNextSound(syllable, Syllables(syllable).IndexOfNuclues - 1)

                    If Nucleus.Contains(PhoneticLength) Then
                        'Long nucleus

                        'Checks that the vowel is a long allophone
                        If Not SwedishLongVowels_IPA.Contains(Nucleus) Then
                            ManualEvaluations.Add(PhoneticMarkingsTypes.SyllableWeightError.ToString & ", Syllable: " & syllable)
                            TotalErrorsFound += 1

                            If LogAllErrors = True Then SendInfoToLog(OrthographicForm & vbTab & String.Join(" ", BuildExtendedIpaArray), "SyllableWeightErrorWords_IncorrectShortVowel")

                        End If


                        'If next sound is a long consonant
                        If SwedishConsonants_IPA.Contains(NextSound) And NextSound.Contains(PhoneticLength) Then
                            ManualEvaluations.Add(PhoneticMarkingsTypes.SyllableWeightError.ToString & ", syllable:  " & syllable)
                            TotalErrorsFound += 1

                            If LogAllErrors = True Then SendInfoToLog(OrthographicForm & vbTab & String.Join(" ", BuildExtendedIpaArray), "SyllableWeightErrorWords_StressedLongLong")

                        End If

                    Else
                        'Checks that the vowel is a short allophone
                        If Not SwedishShortVowels_IPA.Contains(Nucleus) Then
                            ManualEvaluations.Add(PhoneticMarkingsTypes.SyllableWeightError.ToString & ", Syllable: " & syllable)
                            TotalErrorsFound += 1

                            If LogAllErrors = True Then SendInfoToLog(OrthographicForm & vbTab & String.Join(" ", BuildExtendedIpaArray), "SyllableWeightErrorWords_IncorrectReducedLongVowel")

                        End If

                        'Short nucleus
                        If Not (SwedishConsonants_IPA.Contains(NextSound) And NextSound.Contains(PhoneticLength)) Then
                            ManualEvaluations.Add(PhoneticMarkingsTypes.SyllableWeightError.ToString & ", syllable:    " & syllable)
                            TotalErrorsFound += 1

                            If LogAllErrors = True Then SendInfoToLog(OrthographicForm & vbTab & String.Join(" ", BuildExtendedIpaArray), "SyllableWeightErrorWords_StressedShortShort")
                        End If
                    End If

                End If
            End If
        Next

        Return TotalErrorsFound

    End Function

    Public Function MarkPhoneticLengthInWrongPlace() As Integer

        Dim TotalErrors As Integer = 0

        For syllable = 0 To Syllables.Count - 1

            If Syllables(syllable).IsStressed = True Then

                'Getting the nucleus
                Dim Nucleus As String = Syllables(syllable).Phonemes(Syllables(syllable).IndexOfNuclues - 1)
                Dim NextSound As String = Syllables.GetNextSound(syllable, Syllables(syllable).IndexOfNuclues - 1)

                If Nucleus.Contains(PhoneticLength) Then
                    'Long nucleus
                    'The rest of the syllable should be short
                    For ph = Syllables(syllable).IndexOfNuclues To Syllables(syllable).Phonemes.Count - 1
                        If Syllables(syllable).Phonemes(ph).Contains(PhoneticLength) Then

                            'Logging error
                            SendInfoToLog(vbTab & String.Join(" ", BuildExtendedIpaArray), "LengthErrors")
                            TotalErrors += 1
                            ManualEvaluations.Add("LengthError")

                        End If
                    Next

                Else
                    'Short nucleus
                    'The following sound should be long, and the rest of the syllable (if there is any), should contain short sounds
                    If Not NextSound.Contains(PhoneticLength) Then

                        'Logging error
                        SendInfoToLog(vbTab & String.Join(" ", BuildExtendedIpaArray), "LengthErrors")
                        TotalErrors += 1
                        ManualEvaluations.Add("LengthError")

                    End If

                    For ph = Syllables(syllable).IndexOfNuclues + 1 To Syllables(syllable).Phonemes.Count - 1
                        If Syllables(syllable).Phonemes(ph).Contains(PhoneticLength) Then

                            'Logging error
                            SendInfoToLog(vbTab & String.Join(" ", BuildExtendedIpaArray), "LengthErrors")
                            TotalErrors += 1
                            ManualEvaluations.Add("LengthError")

                        End If
                    Next
                End If

            Else

                For ph = 0 To Syllables(syllable).Phonemes.Count - 1

                    'Skipping to next on the first phoneme if the syllable has an ambigous onset
                    If Syllables(syllable).AmbigousOnset = True And ph = 0 Then
                        Continue For
                    End If

                    If Syllables(syllable).Phonemes(ph).Contains(PhoneticLength) Then

                        'Logging error
                        SendInfoToLog(vbTab & String.Join(" ", BuildExtendedIpaArray), "LengthErrors")
                        TotalErrors += 1
                        ManualEvaluations.Add("LengthError")

                    End If
                Next
            End If
        Next

        Return TotalErrors

    End Function


    Public Enum PhoneticMarkingsTypes
        WordWithoutSyllable
        SyllableWithoutNuclues
        SyllableWeightError
        MultipleMainStress
        MultipleSecondaryStress
        NoMainStress
        StressPositionConflict
        NonConsonantInCoda
        NonConsonantInOnset
        NonVowelInNucleus
        Homograph
        Homophone
    End Enum


    ''' <summary>
    ''' Using the IPA characters stored in Syllable.Phonemes together with the vowel list to determine
    ''' index of syllable nucleus (1-based index), length of syllable onset and length of syllable coda.
    ''' Every time DetermineSyllableIndices is run, a testing of syllable structure is also perfored (MarkSyllableStructureErrors)
    ''' </summary>
    ''' <param name="SyllableZeroBaseIndex"></param>
    ''' <returns>Returns the number of errors found.</returns>
    Public Function DetermineSyllableIndices(Optional SyllableZeroBaseIndex As Integer? = Nothing, Optional ByVal SetAlternativeSyllabification As Boolean = False) As Integer

        If SetAlternativeSyllabification = False Then

            For syllable = 0 To Syllables.Count - 1

                'Setting syllable to SyllableZeroBaseIndex if it is the only syllable which should be analysed
                If SyllableZeroBaseIndex IsNot Nothing Then syllable = SyllableZeroBaseIndex

                'Determine syllable indices using the reduced phoneme array (as also the final phoneme array will look like)
                Syllables(syllable).IndexOfNuclues = 0 ' Resetting inputSyllable.indexOfNuclues
                For phoneme = 0 To Syllables(syllable).Phonemes.Count - 1
                    If SwedishVowels_IPA.Contains(Syllables(syllable).Phonemes(phoneme)) Then
                        Syllables(syllable).IndexOfNuclues = phoneme + 1
                        Exit For
                    End If
                Next

                Syllables(syllable).SyllableLength = Syllables(syllable).Phonemes.Count

                If Not Syllables(syllable).IndexOfNuclues = 0 Then
                    Syllables(syllable).LengthOfOnset = Syllables(syllable).IndexOfNuclues - 1
                    Syllables(syllable).LengthOfCoda = Syllables(syllable).SyllableLength - Syllables(syllable).IndexOfNuclues
                End If

                'Aborting loop if only SyllableZeroBaseIndex should be checked
                If SyllableZeroBaseIndex IsNot Nothing Then Exit For

            Next

            'Search the syllables for errors
            Dim TotalErrorsFound As Integer = MarkSyllableStructureErrors()

            Return TotalErrorsFound

        Else
            For syllable = 0 To Syllables_AlternateSyllabification.Count - 1

                'Setting syllable to SyllableZeroBaseIndex if it is the only syllable which should be analysed
                If SyllableZeroBaseIndex IsNot Nothing Then syllable = SyllableZeroBaseIndex

                'Determine syllable indices using the reduced phoneme array (as also the final phoneme array will look like)
                Syllables_AlternateSyllabification(syllable).IndexOfNuclues = 0 ' Resetting inputSyllable.indexOfNuclues
                For phoneme = 0 To Syllables_AlternateSyllabification(syllable).Phonemes.Count - 1
                    If SwedishVowels_IPA.Contains(Syllables_AlternateSyllabification(syllable).Phonemes(phoneme)) Then
                        Syllables_AlternateSyllabification(syllable).IndexOfNuclues = phoneme + 1
                        Exit For
                    End If
                Next

                Syllables_AlternateSyllabification(syllable).SyllableLength = Syllables_AlternateSyllabification(syllable).Phonemes.Count

                If Not Syllables_AlternateSyllabification(syllable).IndexOfNuclues = 0 Then
                    Syllables_AlternateSyllabification(syllable).LengthOfOnset = Syllables_AlternateSyllabification(syllable).IndexOfNuclues - 1
                    Syllables_AlternateSyllabification(syllable).LengthOfCoda = Syllables_AlternateSyllabification(syllable).SyllableLength - Syllables_AlternateSyllabification(syllable).IndexOfNuclues
                End If

                'Aborting loop if only SyllableZeroBaseIndex should be checked
                If SyllableZeroBaseIndex IsNot Nothing Then Exit For

            Next

            'Search the syllables for errors
            Dim TotalErrorsFound As Integer = MarkSyllableStructureErrors()

            Return TotalErrorsFound

        End If

    End Function


    ''' <summary>
    ''' Determines if the syllables with the zero-base index of SyllableZeroBaseIndex is open or closed, or if SyllableZeroBaseIndex is not supplied goes through all syllables in the input word.
    ''' Requires DetectAmbigousCoda to be already run.
    ''' </summary>
    ''' <param name="SyllableZeroBaseIndex"></param>
    Public Sub DetermineSyllableOpenness(Optional SyllableZeroBaseIndex As Integer? = Nothing)

        For syllable = 0 To Syllables.Count - 1

            'Setting syllable to SyllableZeroBaseIndex if it is the only syllable which shouldbe analysed
            If SyllableZeroBaseIndex IsNot Nothing Then syllable = SyllableZeroBaseIndex

            Select Case Syllables(syllable).LengthOfCoda
                Case 0
                    Syllables(syllable).SyllableCodaType = Word.SyllableTypes.Open
                Case 1
                    If Syllables(syllable).Phonemes(Syllables(syllable).Phonemes.Count - 1) = ZeroPhoneme Then
                        Syllables(syllable).SyllableCodaType = Word.SyllableTypes.Open
                    Else
                        Syllables(syllable).SyllableCodaType = Word.SyllableTypes.Closed
                    End If

                Case > 1
                    Syllables(syllable).SyllableCodaType = Word.SyllableTypes.Closed
            End Select

            'Aborting loop if only SyllableZeroBaseIndex shold be checked
            If SyllableZeroBaseIndex IsNot Nothing Then Exit For

        Next

    End Sub

    ''' <summary>
    ''' Detects ambigous syllable boundaries in the word, using the syllables array.
    ''' </summary>
    Public Function DetectAmbigousSyllableBoundaries(Optional ByVal LogWords As Boolean = False) As Boolean

        Dim AmbiguityDetected As Boolean = False


        For syllable = 0 To Syllables.Count - 1

            'If it's not the last syllable
            If syllable <> Syllables.Count - 1 Then

                'If it's the same sound on both sides of the syllable boundary, and this sound is a consonant
                If Syllables(syllable).Phonemes(Syllables(syllable).Phonemes.Count - 1).Trim(PhoneticLength) =
                Syllables(syllable + 1).Phonemes(0).Trim(PhoneticLength) And
                SwedishConsonants_IPA.Contains(Syllables(syllable).Phonemes(Syllables(syllable).Phonemes.Count - 1)) Then

                    Syllables(syllable).AmbigousCoda = True
                    Syllables(syllable + 1).AmbigousOnset = True

                    AmbiguityDetected = True

                    If LogWords = True Then
                        SendInfoToLog(OrthographicForm & vbTab & String.Concat(BuildExtendedIpaArray()) & vbTab &
                                        "Double sound: " & vbTab & Syllables(syllable).Phonemes(Syllables(syllable).Phonemes.Count - 1), "AmbiSyllabicConsonants")
                    End If

                End If

                'If the syllable is stressed, have a short nucleus, and no coda, the coda should be the first consonant in the next syllable, 
                'which means that the syllable boundary can be ambigous

                'If it is established that there is a syllable nucleus
                If Not Syllables(syllable).IndexOfNuclues = 0 Then

                    'If it has no coda
                    If Syllables(syllable).LengthOfCoda = 0 Then

                        'If the syllable is stressed
                        If Syllables(syllable).IsStressed = True Then

                            'If the nucleus is a short vowel
                            If SwedishShortVowels_IncludingLongReduced_IPA.Contains(Syllables(syllable).Phonemes(Syllables(syllable).IndexOfNuclues - 1)) Then

                                'If the following syllable starts with a consonant
                                If SwedishConsonants_IPA.Contains(Syllables(syllable + 1).Phonemes(0)) Then

                                    'Marking as ambisyllabic boundary
                                    Syllables(syllable).AmbigousCoda = True
                                    Syllables(syllable + 1).AmbigousOnset = True

                                    AmbiguityDetected = True

                                    If LogWords = True Then
                                        SendInfoToLog(OrthographicForm & vbTab & String.Concat(BuildExtendedIpaArray()) & vbTab &
                                        "StressedShortNucleusWithoutCoda: " & vbTab & Syllables(syllable).Phonemes(Syllables(syllable).Phonemes.Count - 1), "AmbiSyllabicConsonants")
                                    End If

                                End If
                            End If
                        End If
                    End If
                End If
            End If

        Next

        Return AmbiguityDetected

    End Function


#End Region

#Region "Orthography"

    ''' <summary>
    ''' Sets the property OrthographicFormContainsSpecialCharacter to True if any character in the word is not included in the NormalOrthographicCharacters list.
    ''' </summary>
    ''' <param name="NormalOrthographicCharacters">A list containing a set of normal orthographic characters.</param>
    Public Function MarkSpecialCharacterWords_ByNormalCharList(ByRef NormalOrthographicCharacters As List(Of String), Optional ByRef WordsWithSpecialCharacterCount As Integer = 0) As Boolean

        'Resetting OrthographicFormContainsSpecialCharacter 
        OrthographicFormContainsSpecialCharacter = False

        For character = 0 To OrthographicForm.Count - 1
            If Not NormalOrthographicCharacters.Contains(OrthographicForm.Substring(character, 1)) Then
                OrthographicFormContainsSpecialCharacter = True
                WordsWithSpecialCharacterCount += 1
                Return OrthographicFormContainsSpecialCharacter
            End If
        Next

        Return OrthographicFormContainsSpecialCharacter

    End Function



    Public Sub CountComplexGraphemes(Optional ByRef SetUnresolved_p2g_Character As String = "!", Optional ByVal ExcludeGraphemesWithUnresolved_p2_g_Character As Boolean = True)

        'The function of the parameter SetUnresolved_p2g_Character was added 2017-11-07

        If Sonographs_Letters IsNot Nothing Then

            'Counting number of digraphs
            Dim DigraphCount As Integer = 0
            For Each CurrentGrapheme In Sonographs_Letters
                If CurrentGrapheme.Length = 2 Then
                    If ExcludeGraphemesWithUnresolved_p2_g_Character = True Then
                        If Not CurrentGrapheme.Contains(SetUnresolved_p2g_Character) Then DigraphCount += 1
                    Else
                        DigraphCount += 1
                    End If
                End If
            Next
            Me.DiGraphCount = DigraphCount

            'Counting number of trigraphs
            Dim TrigraphCount As Integer = 0
            For Each CurrentGrapheme In Sonographs_Letters
                If CurrentGrapheme.Length = 3 Then
                    If ExcludeGraphemesWithUnresolved_p2_g_Character = True Then
                        If Not CurrentGrapheme.Contains(SetUnresolved_p2g_Character) Then TrigraphCount += 1
                    Else
                        TrigraphCount += 1
                    End If
                End If
            Next
            Me.TriGraphCount = TrigraphCount

            'Counting number of LongGraphemes (longer than tre letters)
            Dim LongGraphemeCount As Integer = 0
            For Each CurrentGrapheme In Sonographs_Letters
                If CurrentGrapheme.Length > 3 Then
                    If ExcludeGraphemesWithUnresolved_p2_g_Character = True Then
                        If Not CurrentGrapheme.Contains(SetUnresolved_p2g_Character) Then LongGraphemeCount += 1
                    Else
                        LongGraphemeCount += 1
                    End If
                End If
            Next
            Me.LongGraphemesCount = LongGraphemeCount

        End If

    End Sub



#End Region

#Region "WordClassesAndLemmas"


    Public Function WordIsOnlyOneWordClass(ByVal WordClassCode As String) As Boolean

        Dim Result As Boolean = False

        Dim PMCount As Integer = 0
        Dim PoSCount As Integer = 0

        For n = 0 To AllPossiblePoS.Count - 1
            If Not AllPossiblePoS(n).Item1 = "" Then
                PoSCount += 1
                If AllPossiblePoS(n).Item1.StartsWith(WordClassCode) Then
                    PMCount += 1
                End If
            End If
        Next

        If PMCount = PoSCount Then
            Return True
        Else
            Return False
        End If

    End Function



    ''' <summary>
    ''' Determines if a word is assigned to one or more of the word classes in ClassesToLookFor. Returns true a target word class is found, or false if no target word class is found among the possible word class assignment of the word.
    ''' </summary>
    ''' <param name="TargetWordClasses"></param>
    ''' <returns></returns>
    Public Function DetectWordClass(ByVal TargetWordClasses() As String) As Boolean

        Dim HasWordClassData As Boolean = False

        For AWCi = 0 To AllPossiblePoS.Count - 1

            If Not AllPossiblePoS(AWCi).Item1.Trim = "" Then 'Skips if the word class assignment is empty
                For IncWCi = 0 To TargetWordClasses.Count - 1 'Goes through each item in the TargetWordClasses array
                    If AllPossiblePoS(AWCi).Item1.Trim.StartsWith(TargetWordClasses(IncWCi)) Then 'Checks if the current possible word class equals the current item in the TargetWordClasses list
                        Return True 'Returns true if a target word class is detected
                    End If
                Next
            End If
        Next

        Return False 'Returns false if no target word class was detected

    End Function


    ''' <summary>
    ''' Determines if a word has only excluded word classes. Returns true if the word should be excluded (i.e. has only excluded wordclasses), and false if the word should be retained.
    ''' (Returns false if ExcludedWordClasses is left empty, or if no word class is assigned to the word.)
    ''' </summary>
    ''' <param name="ExcludedWordClasses"></param>
    ''' <returns></returns>
    Public Function ExcludeWordDueToWordClasses(ByVal ExcludedWordClasses() As String) As Boolean

        Dim HasWordClassData As Boolean = False

        For AWCi = 0 To AllPossiblePoS.Count - 1
            Dim ExcludedWCCount As Integer = 0 'Detects if the current possible word class is an excluded word class

            If Not AllPossiblePoS(AWCi).Item1.Trim = "" Then 'Skips if the word class assignment is empty
                HasWordClassData = True 'Detects a non-empty word class assignment

                For ExWCi = 0 To ExcludedWordClasses.Count - 1 'Goes through each item in the excluded word class array
                    If AllPossiblePoS(AWCi).Item1.Trim.StartsWith(ExcludedWordClasses(ExWCi)) Then 'Checks if the current possible word class equals the current item in the exclusion list
                        ExcludedWCCount += 1 'Detects if the current possible word class is an excluded word class
                    End If
                Next

                If ExcludedWCCount = 0 Then Return False 'Returns false if the current possible word class was not in the list of excluded word classes (i.e. another (non-empty) word class)

            End If

        Next

        'If the code reached this point, no non-excluded (non-empty) word classes were detected (which means all detected word clases were in the excluded word class list)
        'Returns false if the excluded word class list was empty or if the word did not have any word class assignment (or only empty word class assignments)
        'If not so, returns true to since the word only had word class assignments which all existed in the excluded word class list
        If ExcludedWordClasses.Count = 0 Or HasWordClassData = False Then
            Return False
        Else
            Return True
        End If

    End Function


#End Region

#Region "FrequencyDistributions"


    ''' <summary>
    ''' Calculates Zipf Value, by reading the raw word type frequency already stored in Word.RawWordTypeFrequency.
    ''' The function returns the calculated value and also stores it in Word.ZipfValue_Word
    ''' </summary>
    ''' <param name="CorpusTotalTokenCount">The total number of tokens in the corpus used to set the raw word type frequency.</param>
    ''' <param name="CorpusTotalWordTypeCount">The total number of word types in the corpus used to set the raw word type frequency.</param>
    Public Function CalculateZipfValue_Word(ByRef RawWordTypeFrequency As Long, ByVal CorpusTotalTokenCount As Long, ByVal CorpusTotalWordTypeCount As Integer, Optional ByVal PositionTerm As Integer = 3)

        ZipfValue_Word = CalculateZipfValue(RawWordTypeFrequency, CorpusTotalTokenCount, CorpusTotalWordTypeCount)

        Return ZipfValue_Word

    End Function




#End Region


#Region "Sonographs"


    ''' <summary>
    ''' A finite-state transducer (FST) using the phonetic form of a word, together with a set of phoneme-to-grapheme (p2g) rules to parse the orthographic form of a word into graphemes, where each grapheme corresponds to a phoneme or a combination of phonemes allowed by the rule set.
    ''' </summary>
    ''' <param name="p2g_Settings"></param>
    Public Sub Generate_p2g_Data(ByRef p2g_Settings As p2gParameters)

        'Creates a list to hold attempted spelling segmentations for log purpose
        Dim AttemptedSpellingSegmentations As New List(Of String)
        AttemptedSpellingSegmentations.Add(vbCrLf & vbCrLf & "New Word: " & vbTab & vbTab & OrthographicForm & vbTab & String.Join(" ", BuildExtendedIpaArray))

        'Creating an empty graphemes List
        Sonographs_Letters = New List(Of String)

        Dim EndOfWordIsReached As Boolean = False

        Dim PhoneticForm As List(Of String)
        Dim UseReducedPhoneticForm = False 'Was False during transcription checkings
        If UseReducedPhoneticForm = True Then
            PhoneticForm = BuildReducedIpaArray(False, False)
        Else
            PhoneticForm = BuildExtendedIpaArray(False, True,,,,, False)
        End If

        'Removing stress markers
        Dim PhIndex As Integer = 0
        Do Until PhIndex > PhoneticForm.Count - 1
            If SwedishStressList.Contains(PhoneticForm(PhIndex)) Then
                PhoneticForm.RemoveAt(PhIndex)
            Else
                PhIndex += 1
            End If
        Loop

        'Adding word start marker
        PhoneticForm.Insert(0, p2g_Settings.PhoneToGraphemesDictionary.WordStartMarker)

        'Determining how many word end markers are needed.
        'Adding word final phoneme markers, as many as the length of the longest phoneme combination in the PhoneToGraphemesDictionary + 2, or the highest of allowed unresolved types +2
        'Dim WordEndMarkerCount As Integer = p2g_Settings.PhoneToGraphemesDictionary.MaximumPhonemeCombinationLength
        'If p2g_Settings.MaximumUnresolvedGraphemeJumps > WordEndMarkerCount Then WordEndMarkerCount = p2g_Settings.MaximumUnresolvedGraphemeJumps
        'If p2g_Settings.MaximumUnresolvedPhonemeJumps > WordEndMarkerCount Then WordEndMarkerCount = p2g_Settings.MaximumUnresolvedPhonemeJumps
        'If p2g_Settings.MaximumUnresolvedGraphemeAndPhonemeJumps > WordEndMarkerCount Then WordEndMarkerCount = p2g_Settings.MaximumUnresolvedGraphemeAndPhonemeJumps
        p2g_Settings.PhoneToGraphemesDictionary.NumberOfConcatenatedWordEndMarkers = 1 ' WordEndMarkerCount + 2

        For n = 0 To p2g_Settings.PhoneToGraphemesDictionary.NumberOfConcatenatedWordEndMarkers - 1
            PhoneticForm.Add(p2g_Settings.PhoneToGraphemesDictionary.WordEndMarker)
        Next

        'Creates a copy of PhoneticForm
        Dim PhoneticFormCopy As New List(Of String)
        For Each p In PhoneticForm
            PhoneticFormCopy.Add(p)
        Next

        'Creating a temporary spelling 
        Dim TempSpelling As String = OrthographicForm

        'Checking that the word start and end markers are not in the orthographic form
        If TempSpelling.Contains(p2g_Settings.PhoneToGraphemesDictionary.WordStartMarker) Then
            If p2g_Settings.WarnForIllegalCharactersInSpelling = True Then
                MsgBox("The word " & OrthographicForm & " contains the word start character." & vbTab & vbTab &
                                                    "If you continue, this characther will be temporarily removed during analyis.")
            End If
        End If

        If TempSpelling.Contains(p2g_Settings.PhoneToGraphemesDictionary.WordEndMarker) Then
            If p2g_Settings.WarnForIllegalCharactersInSpelling = True Then
                MsgBox("The word " & OrthographicForm & " contains the word end character." & vbTab & vbTab &
                                                        "If you continue, this characther will be temporarily removed during analyis.")
            End If
        End If

        'Removing any non normalization/junctural characters
        For n = 0 To p2g_Settings.PhoneToGraphemesDictionary.NormalizationCharacters.Count - 1
            TempSpelling = TempSpelling.Replace(p2g_Settings.PhoneToGraphemesDictionary.NormalizationCharacters(n), "")
        Next

        'Adding word start marker
        TempSpelling = p2g_Settings.PhoneToGraphemesDictionary.WordStartMarker & TempSpelling

        'Adding word end marker to the temporary spelling, as many as there are word end markers in the phonetic form
        For n = 0 To p2g_Settings.PhoneToGraphemesDictionary.NumberOfConcatenatedWordEndMarkers - 1
            TempSpelling = TempSpelling & p2g_Settings.PhoneToGraphemesDictionary.WordEndMarker
        Next

        'Replacing any spelling found in TemporarySpellingChangeList
        If p2g_Settings.UseTemporarySpellingChange = True Then
            For line = 0 To p2g_Settings.TemporarySpellingChangeListArray.Length - 1
                Dim LineSplit As String() = p2g_Settings.TemporarySpellingChangeListArray(line).Split(vbTab)
                Dim OriginalSpelling As String = LineSplit(0).Trim
                Dim NewSpelling As String = LineSplit(1).Trim
                TempSpelling = TempSpelling.Replace(OriginalSpelling, NewSpelling)
            Next
        End If

        'Resetting
        Sonographs_Letters.Clear()
        Dim OrthStartIndex As Integer = p2g_Settings.PhoneToGraphemesDictionary.WordStartMarker.Length
        Dim PhoneIndex As Integer = 1
        Dim PhonemeHistory As New List(Of String)

        'Creating an example word phoneme index randomization factor
        Dim rnd As New Random
        Dim ExampleWord_PhonemeIndexRandomization As Single = rnd.Next(0, 1000) / 1000

        'Unresolved p2g:s
        'Setting the number of allowed unresolved p2g, for the 3 different types
        p2g_Settings.MaximumUnresolvedPhonemeJumps = PhoneticForm.Count + p2g_Settings.UnresolvedJumpExtraSteps
        p2g_Settings.MaximumUnresolvedGraphemeJumps = TempSpelling.Length
        'Setting MaximumUnresolvedGraphemeAndPhonemeJumps to the longest of TempSpelling.Length OR  PhoneticForm.Count + p2g_Settings.UnresolvedJumpExtraSteps
        p2g_Settings.MaximumUnresolvedGraphemeAndPhonemeJumps = TempSpelling.Length
        If PhoneticForm.Count + p2g_Settings.UnresolvedJumpExtraSteps > p2g_Settings.MaximumUnresolvedGraphemeAndPhonemeJumps Then p2g_Settings.MaximumUnresolvedGraphemeAndPhonemeJumps = PhoneticForm.Count + p2g_Settings.UnresolvedJumpExtraSteps

        'Creating temporary unresolved p2g variables
        Dim TempAllowUnresolved_p2gs As Boolean = False
        Dim AllowedGraphemeJumps As Integer = 0
        Dim AllowedPhonemeJumps As Integer = 0
        Dim AllowedGraphemeAndPhonemeJumps As Integer = 0
0:

        'Variables to keep the length of sucessful parsing
        Dim LastSuccessPhoneParsingIndex As Integer = 1


        Try

            p2g_CheckNext(p2g_Settings, TempSpelling, OrthStartIndex, PhoneticForm, PhoneIndex,
                              Sonographs_Letters, PhonemeHistory, EndOfWordIsReached, LastSuccessPhoneParsingIndex,
                              False, False, UseReducedPhoneticForm, ExampleWord_PhonemeIndexRandomization,
                              TempAllowUnresolved_p2gs, AllowedGraphemeJumps, AllowedPhonemeJumps, AllowedGraphemeAndPhonemeJumps, -1,
                              AttemptedSpellingSegmentations)


            If EndOfWordIsReached = False Then

                'Resets the phonetic form
                PhoneticForm.Clear()
                For Each p In PhoneticFormCopy
                    PhoneticForm.Add(p)
                Next

                'Activates the jumping function, and updates the number of allowed jumps
                'The jumping funktion allows a number of grapheme or phonemes or both graphemes and phonemes to be unresolved. For each unresolved phonemes/graphems a ! is inserted in the output data (Graphemes/PhonemeBlocks).
                If p2g_Settings.AllowUnresolved_p2gs = True Then
                    TempAllowUnresolved_p2gs = True

                    If AllowedGraphemeJumps < p2g_Settings.MaximumUnresolvedGraphemeJumps And AllowedGraphemeJumps <> -1 Then
                        AllowedGraphemeJumps += 1

                        AttemptedSpellingSegmentations.Add(vbCrLf & vbCrLf & "Increasing AllowedGraphemeJumps to " & AllowedGraphemeJumps)

                        GoTo 0
                    Else
                        'Setting AllowedGraphemeJumps to -1 to disallow this after all allowed grapheme jumps have been tried
                        AllowedGraphemeJumps = -1
                    End If


                    If AllowedGraphemeAndPhonemeJumps < p2g_Settings.MaximumUnresolvedGraphemeAndPhonemeJumps And AllowedGraphemeAndPhonemeJumps <> -1 Then
                        AllowedGraphemeAndPhonemeJumps += 1

                        AttemptedSpellingSegmentations.Add(vbCrLf & vbCrLf & "Increasing AllowedGraphemeAndPhonemeJumps to " & AllowedGraphemeAndPhonemeJumps)

                        GoTo 0
                    Else
                        'Setting AllowedGraphemeAndPhonemeJumps to -1 to disallow this after all allowed Grapheme-And-Phoneme jumps have been tried
                        AllowedGraphemeAndPhonemeJumps = -1
                    End If


                    If AllowedPhonemeJumps < p2g_Settings.MaximumUnresolvedPhonemeJumps And AllowedPhonemeJumps <> -1 Then
                        AllowedPhonemeJumps += 1

                        AttemptedSpellingSegmentations.Add(vbCrLf & vbCrLf & "Increasing AllowedPhonemeJumps to " & AllowedPhonemeJumps)

                        GoTo 0
                    Else
                        'Setting AllowedGraphemeJumps to -1 to disallow this after all allowed phoneme jumps have been tried
                        AllowedPhonemeJumps = -1
                    End If

                End If
            End If


        Catch ex As Exception

            Dim ExceptionString As String = PhoneticForm(LastSuccessPhoneParsingIndex)
            If LastSuccessPhoneParsingIndex - 1 >= 0 Then ExceptionString &= " " & PhoneticForm(LastSuccessPhoneParsingIndex - 1)
            ExceptionString &= " " & String.Join(" ", PhoneticForm, 0, LastSuccessPhoneParsingIndex)
            MsgBox("Graphemes so far " & String.Join(" ", Sonographs_Letters) &
                                       "Last successful phoneme " & PhoneticForm(LastSuccessPhoneParsingIndex) & ex.ToString,, "An error occurred.")
            Exit Sub

        End Try


        'Removing word start markers from the grapheme array
        Dim x As Integer = 0
        Do Until x > Sonographs_Letters.Count - 1
            If Sonographs_Letters(x) = p2g_Settings.PhoneToGraphemesDictionary.WordStartMarker Then
                Sonographs_Letters.RemoveAt(x)
            Else
                x += 1
            End If
        Loop

        'Removing all but word end marker from the grapheme array
        x = 0
        Do Until x > Sonographs_Letters.Count - 1
            If Sonographs_Letters(x) = p2g_Settings.PhoneToGraphemesDictionary.WordEndMarker Then
                Sonographs_Letters.RemoveAt(x)
            Else
                x += 1
            End If
        Loop

        'Removing word start markers from the phoneme history array
        x = 0
        Do Until x > PhonemeHistory.Count - 1
            If PhonemeHistory(x) = p2g_Settings.PhoneToGraphemesDictionary.WordStartMarker Then
                PhonemeHistory.RemoveAt(x)
            Else
                x += 1
            End If
        Loop

        'Removing all but word end marker from the phoneme history array
        x = 0
        Do Until x > PhonemeHistory.Count - 1
            If PhonemeHistory(x) = p2g_Settings.PhoneToGraphemesDictionary.WordEndMarker Then
                PhonemeHistory.RemoveAt(x)
            Else
                x += 1
            End If
        Loop

        'Storing phoneme blocks, that match the grapheme blocks used
        If EndOfWordIsReached = True Then

            'TODO: does word start and word end markers need to be added?
            Sonographs_Pronunciation = PhonemeHistory
        End If

        'Sending words that were parsed by jumping to log for examination
        If EndOfWordIsReached = True And TempAllowUnresolved_p2gs = True Then

            SendInfoToLog(String.Join(" ", PhoneticForm) & vbTab & String.Join("|", Sonographs_Pronunciation) & vbTab & String.Join(" ", Sonographs_Letters) & vbTab & OrthographicForm, "p2g_JumpWords")

        End If


        'Exporting data
        Dim OutputString As String = ""
        If EndOfWordIsReached = False Then

            If p2g_Settings.ErrorDictionary IsNot Nothing Then

                Dim ErrorKey As String = ""
                Dim DictErrorKey As String = ""
                OutputString = ""

                Dim FirstPart As String = ""
                Dim SecondPart As String = ""
                Dim ThirdPart As String = ""

                If LastSuccessPhoneParsingIndex > -1 Then
                    SecondPart = PhoneticForm(LastSuccessPhoneParsingIndex)
                End If
                If LastSuccessPhoneParsingIndex > 0 Then FirstPart = PhoneticForm(LastSuccessPhoneParsingIndex - 1)
                If LastSuccessPhoneParsingIndex + 1 < PhoneticForm.Count Then ThirdPart = PhoneticForm(LastSuccessPhoneParsingIndex + 1)

                ErrorKey = FirstPart & " " & SecondPart & " " & ThirdPart
                DictErrorKey = SecondPart & vbTab & ThirdPart


                OutputString = ErrorKey & " | " & String.Join(" ", PhoneticForm.ToArray, 0, LastSuccessPhoneParsingIndex)

                ManualEvaluations.Add("TsDisagreement " & OutputString)

                'Filling up the error dictionary

                If Not p2g_Settings.ErrorDictionary.ContainsKey(DictErrorKey) Then
                    'Add the key
                    Dim Value As New WordGroup.CountExamples
                    Value.Count = 1
                    Value.Examples.Add(OrthographicForm & " " & TempSpelling & " " & String.Join(" ", PhoneticForm) & " " & String.Concat(BuildReducedIpaArray))

                    'Earlier code: Value.Examples.Add(OrthographicForm & " " & String.Join(" ", PhoneticForm) & " " & String.Concat(BuildExtendedIpaArra()))
                    p2g_Settings.ErrorDictionary.Add(DictErrorKey, Value)
                Else
                    'Increase its value, and add example
                    p2g_Settings.ErrorDictionary(DictErrorKey).Count += 1
                    p2g_Settings.ErrorDictionary(DictErrorKey).Examples.Add(OrthographicForm & " " & TempSpelling & " " & String.Join(" ", PhoneticForm) & " " & String.Concat(BuildReducedIpaArray))
                    'Earlier code: ErrorDictionary(DictErrorKey).Examples.Add(OrthographicForm & " " & String.Join(" ", PhoneticForm) & " " & String.Concat(BuildExtendedIpaArray()))
                End If
            End If


        End If

        'Exporting graphemes
        If p2g_Settings.ListOfGraphemes IsNot Nothing Then
            p2g_Settings.ListOfGraphemes.Add(OrthographicForm & vbTab & String.Join(" ", Sonographs_Letters).Replace(p2g_Settings.PhoneToGraphemesDictionary.WordStartMarker, "").Replace(p2g_Settings.PhoneToGraphemesDictionary.WordEndMarker, "") &
                                                    vbTab & String.Join(" ", PhoneticForm) & vbTab & EndOfWordIsReached.ToString)

        End If

        'Exporting attempted spelling segmentations to file
        If p2g_Settings.ExportAttemptedSpellingSegmentations = True Then
            If p2g_Settings.ExportOnlyFailedAttemptedSpellingSegmentations = False Then
                SendInfoToLog(String.Join(vbCrLf, AttemptedSpellingSegmentations), "AttemptedSpellingSegmentations")
            Else
                If TempAllowUnresolved_p2gs = True Then
                    SendInfoToLog(String.Join(vbCrLf, AttemptedSpellingSegmentations), "AttemptedSpellingSegmentations")
                End If
            End If
        End If


    End Sub

    Private Function p2gTest(ByRef TempSpelling As String, ByRef OrthStartIndex As Integer,
                                 ByRef PhoneticList As List(Of String), ByRef CurrentPhoneArrayIndex As Integer,
                                 ByRef PhonemeLength As Integer, ByRef Grapheme As PhoneToGraphemes.Grapheme,
                                 ByRef AttemptedSpellingSegmentations As List(Of String),
                                 ByRef EndOfWordIsReached As Boolean,
                                 ByRef CurrentGraphemeIndex As Integer,
                                 ByRef GraphemesFound As List(Of String),
                                 ByRef p2g_Settings As p2gParameters) As Boolean


        'Testing conditions for the possible spellings of the phoneme 

        'Testing if the grapheme matches the next grapheme in the orthographic form
        If OrthStartIndex + Grapheme.PossibleSpelling.Length > TempSpelling.Length Then
            Return False 'This means that the grapheme is to long to fit into the word
        Else
            If Not TempSpelling.Substring(OrthStartIndex, Grapheme.PossibleSpelling.Length) = Grapheme.PossibleSpelling Then
                'Aborting if the orthographic form does not contain the current letter
                Return False
            End If
        End If


        'Testing Grapheme pre And post phoneme conditions
        'If both types of conditions exist, both must be fulfilled. If only one condition exists, only that one needs to be fulfilled
        'Testing pre phoneme conditions

        Dim PreAndPostPhonemeConditionApplies As Boolean = False
        Dim PreAndPostPhonemeConditionFulfilled As Boolean = False

        If Grapheme.PreAndPostPhonemeConditions.Count > 0 Then

            PreAndPostPhonemeConditionApplies = True

            'Testing each condition
            For c = 0 To Grapheme.PreAndPostPhonemeConditions.Count - 1

                Dim PrePhonemeConditionApplies As Boolean = False
                Dim PostPhonemeConditionApplies As Boolean = False
                Dim PrePhonemeConditionFulFilled As Boolean = False
                Dim PostPhonemeConditionFulFilled As Boolean = False
                Dim PreTestString As String = ""
                Dim PostTestString As String = ""

                'Parsing the (tab-delimited) conditions
                Dim ConditionSplit() As String = Grapheme.PreAndPostPhonemeConditions(c).Split(vbTab)
                Dim PreCondition As String = ConditionSplit(0)
                Dim PostCondition As String = ConditionSplit(1)

                If PreCondition <> "" Then PrePhonemeConditionApplies = True
                If PostCondition <> "" Then PostPhonemeConditionApplies = True

                If PrePhonemeConditionApplies = True Then

                    Dim PreConditionAlternatives() As String = PreCondition.Split("|")
                    For n = 0 To PreConditionAlternatives.Length - 1
                        Dim CurrentCondition As String = PreConditionAlternatives(n).Trim

                        'Testing first, that we're not outside the word
                        If CurrentPhoneArrayIndex - 1 >= 0 Then
                            If PhoneticList(CurrentPhoneArrayIndex - 1) = CurrentCondition Then
                                PrePhonemeConditionFulFilled = True
                            End If
                        End If
                    Next
                End If

                If PostPhonemeConditionApplies = True Then

                    Dim PostConditionAlternatives() As String = PostCondition.Split("|")
                    For n = 0 To PostConditionAlternatives.Length - 1
                        Dim CurrentCondition As String = PostConditionAlternatives(n).Trim

                        'Testing first, that we're not outside the word
                        If CurrentPhoneArrayIndex + PhonemeLength <= PhoneticList.Count Then
                            If PhoneticList(CurrentPhoneArrayIndex + PhonemeLength) = CurrentCondition Then
                                PostPhonemeConditionFulFilled = True
                            End If
                        End If
                    Next
                End If

                'Testing if conditions are fulfilled
                'If both types of conditions exist, both must be fulfilled. If only one condition exists, only that one needs to be fulfilled
                If PrePhonemeConditionApplies = True And PostPhonemeConditionApplies = True Then
                    If PrePhonemeConditionFulFilled = True And PostPhonemeConditionFulFilled = True Then
                        PreAndPostPhonemeConditionFulfilled = True
                    End If
                End If

                If PrePhonemeConditionApplies = True And PostPhonemeConditionApplies = False Then
                    If PrePhonemeConditionFulFilled = True Then
                        PreAndPostPhonemeConditionFulfilled = True
                    End If
                End If

                If PrePhonemeConditionApplies = False And PostPhonemeConditionApplies = True Then
                    If PostPhonemeConditionFulFilled = True Then
                        PreAndPostPhonemeConditionFulfilled = True
                    End If
                End If
            Next
        End If

        'Aborting if PreAndPostGraphemeCondition Applies and is not fulfilled
        If PreAndPostPhonemeConditionApplies = True And PreAndPostPhonemeConditionFulfilled = False Then Return False



        'Testing grapheme pre and post grapheme conditions
        'If both types of conditions exist, both must be fulfilled. If only one condition exists, only that one needs to be fulfilled
        'Testing pre grapheme conditions

        Dim PreAndPostGraphemeConditionApplies As Boolean = False
        Dim PreAndPostGraphemeConditionFulfilled As Boolean = False

        If Grapheme.PreAndPostGraphemeConditions.Count > 0 Then

            PreAndPostGraphemeConditionApplies = True

            'Testing each condition
            For c = 0 To Grapheme.PreAndPostGraphemeConditions.Count - 1

                Dim PreGraphemeConditionApplies As Boolean = False
                Dim PostGraphemeConditionApplies As Boolean = False
                Dim PreGraphemeConditionFulFilled As Boolean = False
                Dim PostGraphemeConditionFulFilled As Boolean = False
                Dim PreTestString As String = ""
                Dim PostTestString As String = ""

                'Parsing the (tab-delimited) conditions
                Dim ConditionSplit() As String = Grapheme.PreAndPostGraphemeConditions(c).Split(vbTab)
                Dim PreCondition As String = ConditionSplit(0)
                Dim PostCondition As String = ConditionSplit(1)

                If PreCondition <> "" Then PreGraphemeConditionApplies = True
                If PostCondition <> "" Then PostGraphemeConditionApplies = True

                If PreGraphemeConditionApplies = True Then

                    Dim PreConditionAlternatives() As String = PreCondition.Split("|")
                    For n = 0 To PreConditionAlternatives.Length - 1
                        Dim CurrentCondition As String = PreConditionAlternatives(n).Trim
                        If TempSpelling.Substring(0, OrthStartIndex).EndsWith(CurrentCondition) Then
                            PreGraphemeConditionFulFilled = True
                        End If
                    Next

                    'Old code
                    'If TempSpelling.Substring(0, OrthStartIndex).EndsWith(PreCondition) Then
                    'PreGraphemeConditionFulFilled = True
                    'End If
                End If

                If PostGraphemeConditionApplies = True Then

                    Dim PostConditionAlternatives() As String = PostCondition.Split("|")
                    For n = 0 To PostConditionAlternatives.Length - 1
                        Dim CurrentCondition As String = PostConditionAlternatives(n).Trim
                        If TempSpelling.Substring(OrthStartIndex + Grapheme.PossibleSpelling.Length, TempSpelling.Length - (OrthStartIndex + Grapheme.PossibleSpelling.Length)).StartsWith(CurrentCondition) Then
                            PostGraphemeConditionFulFilled = True
                        End If
                    Next

                    'Old code
                    'If TempSpelling.Substring(OrthStartIndex + Grapheme.PossibleSpelling.Length, TempSpelling.Length - (OrthStartIndex + Grapheme.PossibleSpelling.Length)).StartsWith(PostCondition) Then
                    'PostGraphemeConditionFulFilled = True
                    'End If
                End If

                'Testing if conditions are fulfilled
                'If both types of conditions exist, both must be fulfilled. If only one condition exists, only that one needs to be fulfilled
                If PreGraphemeConditionApplies = True And PostGraphemeConditionApplies = True Then
                    If PreGraphemeConditionFulFilled = True And PostGraphemeConditionFulFilled = True Then
                        PreAndPostGraphemeConditionFulfilled = True
                    End If
                End If

                If PreGraphemeConditionApplies = True And PostGraphemeConditionApplies = False Then
                    If PreGraphemeConditionFulFilled = True Then
                        PreAndPostGraphemeConditionFulfilled = True
                    End If
                End If

                If PreGraphemeConditionApplies = False And PostGraphemeConditionApplies = True Then
                    If PostGraphemeConditionFulFilled = True Then
                        PreAndPostGraphemeConditionFulfilled = True
                    End If
                End If
            Next
        End If

        'Aborting if PreAndPostGraphemeCondition Applies and is not fulfilled
        If PreAndPostGraphemeConditionApplies = True And PreAndPostGraphemeConditionFulfilled = False Then Return False

        'Testing that we're not outside the phoneme array
        If CurrentPhoneArrayIndex > PhoneticList.Count - 1 Or OrthStartIndex > TempSpelling.Length - 1 Then
            Return False
        End If

        'Testing if the word end was reached
        If PhoneticList(CurrentPhoneArrayIndex) = p2g_Settings.PhoneToGraphemesDictionary.WordEndMarker And
                                  TempSpelling.Substring(OrthStartIndex, p2g_Settings.PhoneToGraphemesDictionary.WordEndMarker.Length) = p2g_Settings.PhoneToGraphemesDictionary.WordEndMarker Then

            'Adding the last spelling found to AttemptedSpellingSegmentations
            AttemptedSpellingSegmentations.Add(String.Join(" ", GraphemesFound)) ' & vbTab & "S")

            'Setting CurrentGraphemeIndex and exits to previous levels
            EndOfWordIsReached = True
            CurrentGraphemeIndex = GraphemesFound.Count
            Return True
        End If


        Return True 'Returns true if none of the above tests has returned False

    End Function



    Private Sub p2g_CheckNext(ByRef p2g_Settings As p2gParameters,
                                  ByRef Tempspelling As String, ByRef OrthStartIndex As Integer,
                                  ByRef PhoneticForm As List(Of String), ByRef PhoneIndex As Integer,
                                  ByRef GraphemesFound As List(Of String),
                                  ByRef PhonemeHistory As List(Of String),
                                  ByRef EndOfWordIsReached As Boolean,
                                  ByRef LastTrialLastSuccessPhoneParsingIndex As Integer,
                                  ByVal PreviousPhonemeIsPDCSAddition As Boolean,
                                  ByVal PreviousPhonemeIsPDSAddition As Boolean,
                                  ByRef UseReducedPhoneticForm As Boolean,
                                  ByRef ExampleWord_PhonemeIndexRandomization As Single,
                                  ByRef TempAllowUnresolved_p2gs As Boolean,
                                  ByVal AllowedGraphemeJumps As Integer,
                                  ByVal AllowedPhonemeJumps As Integer,
                                  ByVal AllowedGraphemeAndPhonemeJumps As Integer,
                                  ByRef CurrentGraphemeIndex As Integer,
                                  ByRef AttemptedSpellingSegmentations As List(Of String))

        'This sub will generate an array of legal graphemes (that matches the spelling), if such is allowed by the rules in the PhoneToGraphemesDictionary.
        'N.B. There might be other legal grapheme combinations (for example <talljoxe> (if there was such a word!) could be parsed as tal-ljoxe, or as tall-joxe) Sorry for the bad example!

        Try

            'Exits the sub once the end of the word is reached (both phonetic and orthographic form needs to have been ended)
            'If PhoneticForm(PhoneIndex) = p2g_Settings.PhoneToGraphemesDictionary.WordEndMarker And Tempspelling.Substring(OrthStartIndex, p2g_Settings.PhoneToGraphemesDictionary.WordEndMarker.Length) = p2g_Settings.PhoneToGraphemesDictionary.WordEndMarker Then
            'EndOfWordIsReached = True
            'CurrentGraphemeIndex = GraphemesFound.Count
            'End If


            If EndOfWordIsReached = False Then

                Dim SpellingFound As Boolean = False
                Dim CurrentPhonemeString As String = ""

                'Variables for PDCS
                Dim Local_PDCS_Active As Boolean = p2g_Settings.UsePDCS_Addition
                Dim PDCS_Added As Boolean = False
                Dim PDCS_End_PointAdjustment As Integer = 1
                If p2g_Settings.UsePDCS_AdditionOnLastPhoneme = True Then PDCS_End_PointAdjustment = 0
                Dim PDCSInsertionIndex As Integer = 0
                Dim NumberOfAttemptedPDCS As Integer = 0

                'Variables for PDS
                Dim DoSilentGraphemes As Boolean = True
                Dim Local_PDS_Active As Boolean = p2g_Settings.UsePDS_Addition
                Dim NumberOfAttemptedPDS As Integer = 0
                Dim DoSilentPDSGraphemes As Boolean = True 'This is most probably unnessecary
                Dim PDS_Added As Boolean = False
                Dim AllPDSsAttempted As Boolean = False
                Dim PDS_End_PointAdjustment As Integer = 1
                If p2g_Settings.UsePDS_AdditionOnLastPhoneme = True Then PDS_End_PointAdjustment = 0
                Dim PDSInsertionIndex As Integer = 0

                'Variables for phoneme replacements
                Dim Local_PR_Active As Boolean = p2g_Settings.UsePhonemeReplacement
                Dim PR_Performed As Boolean = False
                Dim OriginalPhoneme As String = PhoneticForm(PhoneIndex)
                Dim NumberOfAttemptedRPs As Integer = 0
                Dim Last_PR_Tried As Boolean = False

                'Variables for the AllowUnresolved_p2gs function
                Dim PhonemeJumpCount As Integer = 0
                Dim GraphemeJumpLength As Integer = 0


0: 'Point of restart after failed unresolved_p2g testing

                'Adding unresolved characters
                If GraphemeJumpLength > 0 And PhonemeJumpCount > 0 Then

                    For n = 0 To GraphemeJumpLength - 1 'NB: GraphemeJumpLength should be equal to PhonemeJumpCount, whereby GraphemeJumpLength can be used
                        'Adding the unresolved grapheme
                        Dim CurrentUnresolvedGrapheme As String = Tempspelling.Substring(Tempspelling.Length - 1, 1) 'Using the last character (which would be the word end marker, as default)
                        If OrthStartIndex + n < Tempspelling.Length Then CurrentUnresolvedGrapheme = Tempspelling.Substring(OrthStartIndex + n, 1) 'If we're not after the end of the word, using the appropriate unresolved letter
                        GraphemesFound.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character & CurrentUnresolvedGrapheme)

                        'Adding the unresolved phoneme
                        Dim CurrentUnresolvedPhoneme As String = PhoneticForm(PhoneticForm.Count - 1) 'Using the last phoneme (which would be the word end marker, as default)
                        If PhoneIndex + n < PhoneticForm.Count Then CurrentUnresolvedPhoneme = PhoneticForm(PhoneIndex + n) 'If we're not after the end of the word, using the appropriate unresolved phoneme
                        PhonemeHistory.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character & CurrentUnresolvedPhoneme)

                        'GraphemesFound.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)
                        'PhonemeHistory.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)

                        'Adding the unresolved phoneme
                        'Dim CurrentUnresolvedPhoneme As String = PhoneticForm(PhoneIndex + n)
                        'PhonemeHistory.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character & CurrentUnresolvedPhoneme)
                    Next

                Else
                    For n = 0 To GraphemeJumpLength - 1
                        'GraphemesFound.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)
                        PhonemeHistory.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)

                        'Adding the unresolved grapheme
                        Dim CurrentUnresolvedGrapheme As String = Tempspelling.Substring(Tempspelling.Length - 1, 1) 'Using the last character (which would be the word end marker, as default)
                        If OrthStartIndex + n < Tempspelling.Length Then CurrentUnresolvedGrapheme = Tempspelling.Substring(OrthStartIndex + n, 1) 'If we're not after the end of the word, using the appropriate unresolved letter
                        GraphemesFound.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character & CurrentUnresolvedGrapheme)
                    Next

                    For n = 0 To PhonemeJumpCount - 1
                        GraphemesFound.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)
                        'PhonemeHistory.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)

                        'Adding the unresolved phoneme
                        Dim CurrentUnresolvedPhoneme As String = PhoneticForm(PhoneticForm.Count - 1) 'Using the last phoneme (which would be the word end marker, as default)
                        If PhoneIndex + n < PhoneticForm.Count Then CurrentUnresolvedPhoneme = PhoneticForm(PhoneIndex + n) 'If we're not after the end of the word, using the appropriate unresolved phoneme
                        PhonemeHistory.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character & CurrentUnresolvedPhoneme)
                    Next
                End If


                'If GraphemeJumpLength = PhonemeJumpCount And GraphemeJumpLength > 0 Then
                'For n = 0 To GraphemeJumpLength - 1
                'GraphemesFound.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)
                'PhonemeHistory.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)
                'Next
                'ElseIf GraphemeJumpLength > 0 Then
                'For n = 0 To GraphemeJumpLength - 1
                'PhonemeHistory.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)

                'Adding the unresolved grapheme
                'Dim CurrentUnresolvedGrapheme As String = Tempspelling.Substring(OrthStartIndex + n, 1)
                'GraphemesFound.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character & CurrentUnresolvedGrapheme)
                'Next
                'ElseIf PhonemeJumpCount > 0 Then
                'For n = 0 To PhonemeJumpCount - 1
                'GraphemesFound.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)

                'Adding the unresolved phoneme
                'Dim CurrentUnresolvedPhoneme As String = PhoneticForm(PhoneIndex + n)
                'PhonemeHistory.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character & CurrentUnresolvedPhoneme)

                'Next
                'End If


                'Testing all phoneme lengths in the dictionary, in the reversed order
                'Storing a history of successful p2g conversions graphemes for the phoneme position
                Dim HistoryOfSuccessfulGraphemes As New List(Of String)

                For ReversedPhonemeLength = 0 To p2g_Settings.PhoneToGraphemesDictionary.MaximumPhonemeCombinationLength
                    Dim PhonemeLength As Integer = p2g_Settings.PhoneToGraphemesDictionary.MaximumPhonemeCombinationLength - ReversedPhonemeLength

                    If EndOfWordIsReached = True Then Exit Sub 'GoTo 1

                    'Checking that PhoneIndex + SkippedPhonemeCount + PhonemeLength does not go outside the PhoneticForm array
                    If PhoneIndex + PhonemeJumpCount + PhonemeLength > PhoneticForm.Count Then
                        'Skips if we're outside the PhoneticForm array length
                        Continue For
                    End If


                    Select Case p2g_Settings.LengthSensitiveLookUpPhonemes
                        Case True
                            CurrentPhonemeString = String.Join(" ", PhoneticForm.ToArray, PhoneIndex + PhonemeJumpCount, PhonemeLength).Trim
                        Case False
                            CurrentPhonemeString = String.Join(" ", PhoneticForm.ToArray, PhoneIndex + PhonemeJumpCount, PhonemeLength).Replace(PhoneticLength, "").Trim
                    End Select


                    'Checking that the current phoneme combination exists in the PhoneToGraphemesDictionary
                    If PhonemeLength = 0 Then CurrentPhonemeString = "∅"
                    If Not p2g_Settings.PhoneToGraphemesDictionary.ContainsKey(CurrentPhonemeString) Then Continue For

                    For Each Grapheme In p2g_Settings.PhoneToGraphemesDictionary(CurrentPhonemeString)

                        'Blocking current silent graphemes (with a phoneme length of 0) if a non silent version has already been approved,
                        'Only done the second time and if Temp_UsePDS_Addition is true
                        If PhonemeLength = 0 Then
                            If DoSilentGraphemes = False Then
                                If HistoryOfSuccessfulGraphemes.Contains(Grapheme.PossibleSpelling) Then Continue For
                            End If
                        End If

                        'Checking that the current possible spelling does not go outside the TempSpelling string
                        If Not OrthStartIndex + GraphemeJumpLength + Grapheme.PossibleSpelling.Length <= Tempspelling.Length Then
                            'Skips if we're outside the Tempspelling string length
                            Continue For
                        End If

                        'Adding unresolved graphemes
                        'For n = 0 To SkippedGraphemesLength - 1
                        'GraphemesFound.Add("?")
                        'Next


                        'Testing if the current grapheme matches the orthographic form
                        If Not p2gTest(Tempspelling, OrthStartIndex + GraphemeJumpLength, PhoneticForm,
                                           PhoneIndex + PhonemeJumpCount, PhonemeLength, Grapheme,
                                           AttemptedSpellingSegmentations, EndOfWordIsReached, CurrentGraphemeIndex, GraphemesFound, p2g_Settings) = True Then

                            'Removing unresolved graphemes
                            'For n = 0 To SkippedGraphemesLength - 1
                            'GraphemesFound.RemoveAt(GraphemesFound.Count - 1)
                            'Next

                            'Storing Attempted spelling segmentations
                            'AttemptedSpellingSegmentations.Add(String.Join(" ", GraphemesFound) & " " & Grapheme.PossibleSpelling & vbTab & "F")

                            Continue For
                        Else

                            'SendInfoToLog("AllowedPhonemeJumps:" & AllowedPhonemeJumps & vbCrLf &
                            '    "AllowedGraphemeJumps:" & AllowedGraphemeJumps & vbCrLf &
                            '    "AllowedGraphemeAndPhonemeJumps:" & AllowedGraphemeAndPhonemeJumps & vbCrLf &
                            'vbTab & vbTab & vbTab & vbTab & String.Join(" ", GraphemesFound) & " " & Grapheme.PossibleSpelling & vbTab & String.Join(" ", PhonemeHistory))

                            If EndOfWordIsReached = True Then
                                Exit Sub
                            End If

                            SpellingFound = True

                            'Adding any skipped (unresolved) graphemes/phoneme (Equal number of unresolved characters need to be added, in order for the spelling regularity calculation to word)
                            'If SkippedGraphemesLength > SkippedPhonemeCount Then
                            'For n = 0 To SkippedGraphemesLength - 1
                            'GraphemesFound.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)
                            'PhonemeHistory.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)
                            'Next
                            'Else
                            'For n = 0 To SkippedPhonemeCount - 1
                            'GraphemesFound.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)
                            'PhonemeHistory.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)
                            'Next
                            'End If

                            'If SkippedGraphemesLength = SkippedPhonemeCount And SkippedGraphemesLength > 0 Then
                            'For n = 0 To SkippedGraphemesLength - 1
                            'GraphemesFound.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)
                            'PhonemeHistory.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)
                            'Next
                            'ElseIf SkippedGraphemesLength > SkippedPhonemeCount And SkippedGraphemesLength > 0 Then
                            'For n = 0 To SkippedPhonemeCount - 1
                            'PhonemeHistory.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)
                            'Next
                            'ElseIf SkippedPhonemeCount > SkippedGraphemesLength And SkippedPhonemeCount > 0 Then
                            'For n = 0 To SkippedGraphemesLength - 1
                            'GraphemesFound.Add(p2g_Settings.PhoneToGraphemesDictionary.Unresolved_p2g_Character)
                            'Next
                            'End If

                            'Adding the grapheme found, and calculating spelling commonality/regularity data
                            GraphemesFound.Add(Grapheme.PossibleSpelling)
                            PhonemeHistory.Add(CurrentPhonemeString)
                            HistoryOfSuccessfulGraphemes.Add(Grapheme.PossibleSpelling)

                            'Storing Attempted spelling segmentations
                            AttemptedSpellingSegmentations.Add(String.Join(" ", GraphemesFound)) ' & vbTab & "S")

                            'Noting sucessful parsing length (noted only if the current transition is not an unparsable type)
                            If Not GraphemeJumpLength < 1 And PhonemeJumpCount < 1 Then
                                If PhoneIndex >= LastTrialLastSuccessPhoneParsingIndex Then
                                    LastTrialLastSuccessPhoneParsingIndex = PhoneIndex + PhonemeJumpCount
                                End If
                            End If

                            'Increasing position holders for next phoneme
                            OrthStartIndex += Grapheme.PossibleSpelling.Length + GraphemeJumpLength
                            PhoneIndex += PhonemeLength + PhonemeJumpCount


                            'Checking the next phoneme
                            p2g_CheckNext(p2g_Settings, Tempspelling, OrthStartIndex, PhoneticForm, PhoneIndex,
                                              GraphemesFound, PhonemeHistory, EndOfWordIsReached,
                                              LastTrialLastSuccessPhoneParsingIndex,
                                              PDCS_Added, PDS_Added, UseReducedPhoneticForm, ExampleWord_PhonemeIndexRandomization,
                                              TempAllowUnresolved_p2gs, AllowedGraphemeJumps,
                                              AllowedPhonemeJumps, AllowedGraphemeAndPhonemeJumps, CurrentGraphemeIndex, AttemptedSpellingSegmentations)

                            'If the code comes back here, it either means that a spelling could not be found for the next phoneme, or that the end of the word has been reached
                            'Exiting if end of word is reached

                            'Rolling back start indices to the current phoneme/grapheme start position (since the attempted spelling was incorrect / or there were no more letters (i.e. word end reached))
                            PhoneIndex -= (PhonemeLength + PhonemeJumpCount)
                            OrthStartIndex -= (Grapheme.PossibleSpelling.Length + GraphemeJumpLength)


                            If EndOfWordIsReached = True Then

                                'Rolling back the CurrentGraphemeIndex one step
                                CurrentGraphemeIndex -= 1

                                'Also counting occurences of each phoneme
                                Grapheme.SimpleCount += 1

                                'Counting the occurence of each grapheme
                                p2g_Settings.PhoneToGraphemesDictionary(CurrentPhonemeString).SimpleCount += 1

                                'Storing example words (Only stored if the current example does not contain any unresolved g2p correspondences)
                                If p2g_Settings.UseExampleWordSampling = True And TempAllowUnresolved_p2gs = False Then

                                    p2g_MarkAndExportExampleWords(Tempspelling, PhoneticForm, PhoneIndex,
                      p2g_Settings.PhoneToGraphemesDictionary, PDCS_Added,
                      PDS_Added, PR_Performed,
                      OriginalPhoneme, UseReducedPhoneticForm, ExampleWord_PhonemeIndexRandomization,
                      p2g_Settings.MaximumNumberOfExampleWords, CurrentPhonemeString, Grapheme,
                      PhonemeLength, p2g_Settings.OutputFolder)

                                    'Not exporting abbreviations or foregin words
                                    If Abbreviation = False And ForeignWord = False Then
                                        p2g_MarkAndExportMostCommonExampleWords(Tempspelling, PhoneticForm,
                                                                      p2g_Settings.PhoneToGraphemesDictionary, PDCS_Added,
                                                                      PDS_Added, PR_Performed,
                                                                      OriginalPhoneme, UseReducedPhoneticForm, CurrentPhonemeString, Grapheme,
                                                                      PhonemeLength, p2g_Settings.OutputFolder, PhonemeHistory)
                                    End If
                                End If

                                'Exiting back to the previous phoneme function
                                Exit Sub

                            Else

                                'Removing incorrectly stored graphemes and phonemes
                                GraphemesFound.RemoveAt(GraphemesFound.Count - 1)
                                PhonemeHistory.RemoveAt(PhonemeHistory.Count - 1)

                                'Removing also any skipped/unresolved graphemes and phonemes
                                'Adding any skipped (unresolved) graphemes/phoneme (Equal number of unresolved characters need to be added, in order for the spelling regularity calculation to work)
                                'If SkippedGraphemesLength > SkippedPhonemeCount Then
                                'For n = 0 To SkippedGraphemesLength - 1
                                'GraphemesFound.RemoveAt(GraphemesFound.Count - 1)
                                'PhonemeHistory.RemoveAt(PhonemeHistory.Count - 1)
                                'Next
                                'Else
                                'For n = 0 To SkippedPhonemeCount - 1
                                'GraphemesFound.RemoveAt(GraphemesFound.Count - 1)
                                'PhonemeHistory.RemoveAt(PhonemeHistory.Count - 1)
                                'Next
                                'End If

                                'If SkippedGraphemesLength = SkippedPhonemeCount And SkippedGraphemesLength > 0 Then
                                'For n = 0 To SkippedGraphemesLength - 1
                                'GraphemesFound.RemoveAt(GraphemesFound.Count - 1)
                                'PhonemeHistory.RemoveAt(PhonemeHistory.Count - 1)
                                'Next
                                'ElseIf SkippedGraphemesLength > SkippedPhonemeCount And SkippedGraphemesLength > 0 Then
                                'For n = 0 To SkippedPhonemeCount - 1
                                'PhonemeHistory.RemoveAt(PhonemeHistory.Count - 1)
                                'Next
                                'ElseIf SkippedPhonemeCount > SkippedGraphemesLength And SkippedPhonemeCount > 0 Then
                                'For n = 0 To SkippedGraphemesLength - 1
                                'GraphemesFound.RemoveAt(GraphemesFound.Count - 1)
                                'Next
                                'End If

                            End If
                        End If
                    Next
                Next

                'Removing any added unresolved character
                'Adding unresolved characters
                If GraphemeJumpLength > 0 And PhonemeJumpCount > 0 Then
                    For n = 0 To GraphemeJumpLength - 1 'NB: GraphemeJumpLength should be equal to PhonemeJumpCount, whereby GraphemeJumpLength can be used
                        PhonemeHistory.RemoveAt(PhonemeHistory.Count - 1)
                        GraphemesFound.RemoveAt(GraphemesFound.Count - 1)
                    Next
                Else
                    For n = 0 To GraphemeJumpLength - 1
                        PhonemeHistory.RemoveAt(PhonemeHistory.Count - 1)
                        GraphemesFound.RemoveAt(GraphemesFound.Count - 1)
                    Next
                    For n = 0 To PhonemeJumpCount - 1
                        PhonemeHistory.RemoveAt(PhonemeHistory.Count - 1)
                        GraphemesFound.RemoveAt(GraphemesFound.Count - 1)
                    Next
                End If



                'Testing insertion of a PDCS from next position up to PhoneToGraphemesDictionary.MaximumPhonemeCombinationLength positions after the current phoneme
                'NB PDCS is never allowed on the first phoneme of a word

                If Local_PDCS_Active = True And EndOfWordIsReached = False And PreviousPhonemeIsPDCSAddition = False And PreviousPhonemeIsPDSAddition = False And PhoneIndex > 1 Then

                    If PDCS_Added = True Then

                        'Removes the PDCS that was not working
                        PhoneticForm.RemoveAt(PDCSInsertionIndex)

                        'Resetting PDCS_Added
                        PDCS_Added = False

                    End If

1: 'Point of restart when no PDCS is added due to the PossibleDeletedCompundedSegments of the tested segment
                    If NumberOfAttemptedPDCS < p2g_Settings.PhoneToGraphemesDictionary.MaximumPhonemeCombinationLength Then

                        'Testing if we're on the last phoneme, and skips PDCS if UsePDCS_AdditionOnLastPhoneme = False
                        If PhoneIndex + NumberOfAttemptedPDCS < PhoneticForm.Count - p2g_Settings.PhoneToGraphemesDictionary.NumberOfConcatenatedWordEndMarkers + 1 - PDCS_End_PointAdjustment Then

                            'Testing if the preceding phoneme allows PDCS. If the tested phoneme does not exist as a unique phoneme (which would happen if it is only part of a phoneme combination), 
                            'it could possibly be lacking in the dictionary. Therefore it is first checked if it is present in the dictionary. If not, PDCS is disallowed.
                            '(The reason for checking PDCS on not just the current phoneme is that phoneme combinations e.g. [k s] repressenting <x> (and similar situations) 
                            'needs the [s] to be repeated in order to parse <xs>, and thus PDCS needs to be attempted also further into the word.

                            Dim AllowPDCS As Boolean = False
                            Dim PDCS_PhonemeString As String = PhoneticForm(PhoneIndex + NumberOfAttemptedPDCS - 1)
                            Select Case p2g_Settings.LengthSensitiveLookUpPhonemes
                                Case False
                                    If p2g_Settings.PhoneToGraphemesDictionary.ContainsKey(PDCS_PhonemeString.Replace(PhoneticLength, "")) Then
                                        AllowPDCS = p2g_Settings.PhoneToGraphemesDictionary(PDCS_PhonemeString.Replace(PhoneticLength, "")).PossibleDeletedCompundedSegments
                                    Else
                                        AllowPDCS = False
                                    End If

                                Case True
                                    If p2g_Settings.PhoneToGraphemesDictionary.ContainsKey(PDCS_PhonemeString) Then
                                        AllowPDCS = p2g_Settings.PhoneToGraphemesDictionary(PDCS_PhonemeString).PossibleDeletedCompundedSegments
                                    Else
                                        AllowPDCS = False
                                    End If
                            End Select
                            If AllowPDCS = True Then

                                'Adding a PDCS on the current PhoneticForm index 
                                PDCSInsertionIndex = PhoneIndex + NumberOfAttemptedPDCS
                                Select Case p2g_Settings.UsePDCS_AdditionsWithoutLength
                                    Case False
                                        PhoneticForm.Insert(PDCSInsertionIndex, PDCS_PhonemeString)
                                    Case True
                                        PhoneticForm.Insert(PDCSInsertionIndex, PDCS_PhonemeString.Replace(PhoneticLength, ""))
                                End Select

                                'Increasing NumberOfAttemptedPDCS
                                NumberOfAttemptedPDCS += 1

                                'Noting PDCS as added
                                PDCS_Added = True

                                'Restarting analysis with added PDCS
                                GoTo 0
                            Else

                                'Increasing NumberOfAttemptedPDCS
                                NumberOfAttemptedPDCS += 1
                                GoTo 1

                            End If
                        Else
                            'Turning off PDCS for the current phoneme
                            Local_PDCS_Active = False

                        End If
                    Else
                        'Turning off PDCS for the current phoneme
                        Local_PDCS_Active = False

                    End If
                End If


                'Testing insertion of PDSs after the current phoneme, only if the previous phoneme is not a PDS
                'NB PDCS is never allowed on the first phoneme of a word
                If Local_PDS_Active = True And EndOfWordIsReached = False And PreviousPhonemeIsPDCSAddition = False And PreviousPhonemeIsPDSAddition = False And PhoneIndex > 1 Then

                    'Removes any previously added PDS that was not working, if it has been tried both without and with silent graphemes
                    If PDS_Added = True And DoSilentPDSGraphemes = True Then

                        PhoneticForm.RemoveAt(PDCSInsertionIndex)

                        'Resetting PDCS_Added
                        PDS_Added = False

                    End If


                    'Adding a PDS
                    If Not AllPDSsAttempted = True Then

                        'Testing if we're on the last phoneme, and skips PDS if UsePDS_AdditionOnLastPhoneme = False
                        If PhoneIndex < PhoneticForm.Count - p2g_Settings.PhoneToGraphemesDictionary.NumberOfConcatenatedWordEndMarkers + 1 - PDS_End_PointAdjustment Then

                            If NumberOfAttemptedPDS < p2g_Settings.PhoneToGraphemesDictionary.ListOfPDSs.Count Then

                                'Each PDS phoneme is first run without, and then with silent phonemes
                                If DoSilentPDSGraphemes = True Then
                                    DoSilentGraphemes = False
                                    DoSilentPDSGraphemes = False

                                    'Adding a PDS after the current phoneme
                                    PDCSInsertionIndex = PhoneIndex
                                    PhoneticForm.Insert(PDCSInsertionIndex, p2g_Settings.PhoneToGraphemesDictionary.ListOfPDSs(NumberOfAttemptedPDS))

                                    'Noting PDS as added
                                    PDS_Added = True

                                Else
                                    DoSilentGraphemes = True
                                    DoSilentPDSGraphemes = True

                                    'Going to the next PDS
                                    NumberOfAttemptedPDS += 1

                                    If NumberOfAttemptedPDS = p2g_Settings.PhoneToGraphemesDictionary.ListOfPDSs.Count Then
                                        'Notes that att possible PDSs on this index is attempted
                                        AllPDSsAttempted = True
                                    End If

                                End If

                                'Restarting analysis with added PDCS
                                GoTo 0

                            Else
                                'Turning off PDCS for the current phoneme
                                Local_PDCS_Active = False

                            End If

                        Else
                            'Turning off PDCS for the current phoneme
                            Local_PDCS_Active = False

                        End If
                    End If
                End If


                'Resetting the value of DoSilentGraphemes before further processing
                DoSilentGraphemes = True

                'Testing phoneme replacements
                If Local_PR_Active = True And EndOfWordIsReached = False And PreviousPhonemeIsPDCSAddition = False And PreviousPhonemeIsPDSAddition = False Then

                    'Removes any previously non working phoneme replacements
                    If PR_Performed = True Then

                        PhoneticForm(PhoneIndex) = OriginalPhoneme

                        'Resetting PR_Performed
                        PR_Performed = False

                    End If

                    'Phoneme replacement section
                    If Not Last_PR_Tried = True Then



                        'Getting the replacement phonemes for the current phoneme
                        Dim ListOfReplaceMentPhonemes As New List(Of String)
                        Select Case p2g_Settings.LengthSensitiveLookUpPhonemes
                            Case False
                                ListOfReplaceMentPhonemes = p2g_Settings.PhoneToGraphemesDictionary(OriginalPhoneme.Replace(PhoneticLength, "")).PossibleReplacementPhonemes
                            Case True
                                ListOfReplaceMentPhonemes = p2g_Settings.PhoneToGraphemesDictionary(OriginalPhoneme).PossibleReplacementPhonemes
                        End Select

                        If NumberOfAttemptedRPs < ListOfReplaceMentPhonemes.Count Then

                            'Doing phoneme replacement to the current phoneme index
                            Select Case p2g_Settings.UsePDCS_AdditionsWithoutLength
                                Case False
                                    PhoneticForm(PhoneIndex) = ListOfReplaceMentPhonemes(NumberOfAttemptedRPs)
                                Case True
                                    PhoneticForm(PhoneIndex) = ListOfReplaceMentPhonemes(NumberOfAttemptedRPs).Replace(PhoneticLength, "")
                            End Select

                            'Noting that phoneme replacement is performed
                            PR_Performed = True

                            'Increasing NumberOfAttemptedRPs
                            NumberOfAttemptedRPs += 1

                            'Checking if this was the last available replacement phoneme 
                            If NumberOfAttemptedRPs = ListOfReplaceMentPhonemes.Count Then
                                Last_PR_Tried = True
                            End If

                            'Restarting analysis using the replaced phoneme
                            GoTo 0
                        End If


                    Else
                        'Inactivating phoneme replacement for the current phoneme
                        Local_PR_Active = False
                    End If

                End If


                'Allowing a number of unresolved p2g correspondences
                'This function is activated only after the last phoneme adjustment function is skipped

                If TempAllowUnresolved_p2gs = True And EndOfWordIsReached = False Then

                    If AllowedGraphemeAndPhonemeJumps > 0 Then

                        If GraphemeJumpLength < AllowedGraphemeAndPhonemeJumps And PhonemeJumpCount < AllowedGraphemeAndPhonemeJumps Then

                            'Checking if skipping one extra grapheme takes us outside the end of the string
                            If OrthStartIndex + GraphemeJumpLength + 1 > Tempspelling.Length - 1 And PhoneIndex + PhonemeJumpCount + 1 > PhoneticForm.Count - 1 Then
                                'Not increasing GraphemeJumpLength And PhonemeJumpCount more since it will go outside either Tempspelling Or PhoneticForm
                                GraphemeJumpLength = 0
                                PhonemeJumpCount = 0
                            Else

                                'Skipping one extra grapheme
                                GraphemeJumpLength += 1

                                'Skipping one extra phoneme
                                PhonemeJumpCount += 1

                                'Restarting analysis with skipped phonemes/graphemes
                                GoTo 0
                            End If
                        End If
                    End If
                End If

                If TempAllowUnresolved_p2gs = True And EndOfWordIsReached = False Then
                    If AllowedGraphemeJumps > 0 Then

                        If GraphemeJumpLength < AllowedGraphemeJumps Then

                            'Checking if skipping one extra grapheme takes us outside the end of the string
                            If OrthStartIndex + GraphemeJumpLength + 1 > Tempspelling.Length - 1 Then
                                'Not increasing GraphemeJumpLength more since it will go outside Tempspelling 
                                GraphemeJumpLength = 0
                            Else

                                'Skipping one extra grapheme
                                GraphemeJumpLength += 1

                                'Restarting analysis with skipped phonemes/graphemes
                                GoTo 0
                            End If

                        End If
                    End If
                End If

                If TempAllowUnresolved_p2gs = True And EndOfWordIsReached = False Then

                    If AllowedPhonemeJumps > 0 Then

                        If PhonemeJumpCount < AllowedPhonemeJumps Then

                            'Checking if skipping one extra grapheme takes us outside the end of the string
                            If PhoneIndex + PhonemeJumpCount + 1 > PhoneticForm.Count - 1 Then
                                'Not increasing PhonemeJumpCount more since it will go outside PhoneticForm
                                PhonemeJumpCount = 0
                            Else

                                'Skipping one extra phoneme
                                PhonemeJumpCount += 1

                                'Restarting analysis with skipped phonemes/graphemes
                                GoTo 0
                            End If
                        End If
                    End If
                End If


            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub


    Private Sub p2g_MarkAndExportExampleWords(ByRef Tempspelling As String,
                                                  ByRef PhoneticForm As List(Of String),
                                                  ByRef PhoneIndex As Integer,
                                                  ByRef PhoneToGraphemesDictionary As PhoneToGraphemes,
                                                  ByRef PDCS_Grapheme As Boolean,
                                                  ByRef PDS_Grapheme As Boolean,
                                                  ByRef PR_Grapheme As Integer,
                                                  ByRef OriginalPhoneme As String,
                                                  ByRef UseReducedPhoneticForm As Boolean,
                                                  ByRef ExampleWord_PhonemeIndexRandomization As Single,
                                                  ByRef MaximumNumberOfExampleWords As Integer,
                                                  ByRef CurrentPhonemeString As String,
                                                  ByRef Grapheme As PhoneToGraphemes.Grapheme,
                                                  ByRef PhonemeLength As Integer,
                                                  ByRef OutputFolder As String)

        Try


            'The example words are exported

            'Adding example words if the sample point is active
            'Adding the phoneme only if it repressents the current randomization factor (this means that only one phoneme can be added from the same word)
            'However, before using the randomization function the list is filled using all available examples (so that no grapheme is left empty which has examples)
            Dim RandomizedPhonemeIndex As Integer = Int((PhoneticForm.Count - 1 - 1 - PhoneToGraphemesDictionary.MaximumPhonemeCombinationLength) * ExampleWord_PhonemeIndexRandomization)

            Dim OutputTranscription As String = ""
            If UseReducedPhoneticForm = True Then
                OutputTranscription = String.Concat(BuildReducedIpaArray(False, False))
            Else
                OutputTranscription = String.Concat(BuildExtendedIpaArray(False, True,,,,, False))
            End If

            Dim OutputSpelling As String = Tempspelling.Replace(PhoneToGraphemesDictionary.WordEndMarker, "").Replace(PhoneToGraphemesDictionary.WordStartMarker, "")
            Dim PhoneticFormConcat As String = " [" & String.Concat(PhoneticForm) & "]".Replace(PhoneToGraphemesDictionary.WordEndMarker, "").Replace(PhoneToGraphemesDictionary.WordStartMarker, "")

            If PDCS_Grapheme = True Then

                'Marking the word as containing PDCS addition
                ManualEvaluations.Add("PDCS of " & CurrentPhonemeString)

                'Also exporting these to log file for evaluation
                SendInfoToLog(OrthographicForm & vbTab & OutputTranscription & vbTab & OutputSpelling & PhoneticFormConcat,
                              "PDCS_words", OutputFolder)

                'before using the randomization function the list is filled using all available examples
                If Grapheme.ExamplesWithPDCS.Count < MaximumNumberOfExampleWords Then
                    Grapheme.ExamplesWithPDCS.Add(OrthographicForm & PhoneticFormConcat & " " & OutputTranscription)
                Else

                    'Testing if RandomizedPhonemeIndex is within the index range of the current phoneme combination
                    If RandomizedPhonemeIndex >= PhoneIndex And RandomizedPhonemeIndex < PhoneIndex + PhonemeLength Then

                        'Or else continues filling the list using the randomization function
                        If Grapheme.ExampleSamplingWithPDCSActive = True Then
                            Grapheme.ExamplesWithPDCS.Add(OrthographicForm & PhoneticFormConcat & " " & OutputTranscription)
                            Grapheme.ExampleSamplingWithPDCSActive = False
                            Grapheme.ExamplesWithPDCS.RemoveAt(1)
                        End If
                    End If
                End If
            Else
                'before using the randomization function the list is filled using all available examples
                If Grapheme.ExamplesWithoutPDCS.Count < MaximumNumberOfExampleWords Then
                    Grapheme.ExamplesWithoutPDCS.Add(OrthographicForm & PhoneticFormConcat & " " & OutputTranscription)
                Else
                    'Testing if RandomizedPhonemeIndex is within the index range of the current phoneme combination
                    If RandomizedPhonemeIndex >= PhoneIndex And RandomizedPhonemeIndex < PhoneIndex + PhonemeLength Then
                        'Or else continues filling the list using the randomization function
                        If Grapheme.ExampleSamplingWithoutPDCSActive = True Then
                            Grapheme.ExamplesWithoutPDCS.Add(OrthographicForm & PhoneticFormConcat & " " & OutputTranscription)
                            Grapheme.ExampleSamplingWithoutPDCSActive = False
                            Grapheme.ExamplesWithoutPDCS.RemoveAt(1)
                        End If
                    End If
                End If
            End If

            'Adding PDS exmaples
            If PDS_Grapheme = True Then

                'before using the randomization function the list is filled using all available examples
                If Grapheme.ExamplesWithPDS.Count < MaximumNumberOfExampleWords Then
                    Grapheme.ExamplesWithPDS.Add(OrthographicForm & PhoneticFormConcat & " " & OutputTranscription)
                Else
                    'Testing if RandomizedPhonemeIndex is within the index range of the current phoneme combination
                    If RandomizedPhonemeIndex >= PhoneIndex And RandomizedPhonemeIndex < PhoneIndex + PhonemeLength Then
                        'Or else continues filling the list using the randomization function
                        If Grapheme.ExampleSamplingWithPDSActive = True Then
                            Grapheme.ExamplesWithPDS.Add(OrthographicForm & PhoneticFormConcat & " " & OutputTranscription)
                            Grapheme.ExampleSamplingWithPDSActive = False
                            Grapheme.ExamplesWithPDS.RemoveAt(1)
                        End If
                    End If
                End If

                'Marking the word as containing PDS addition
                ManualEvaluations.Add("PDS of " & CurrentPhonemeString)

                'Also exporting these to log file for evaluation
                SendInfoToLog(OrthographicForm & vbTab & OutputTranscription & vbTab & OutputSpelling & PhoneticFormConcat,
                              "PDS_words", OutputFolder)

            End If

            'Adding Phoneme replacement examples
            If PR_Grapheme = True Then

                'Marking the word as corrected transcription
                CorrectedTranscription = True

                'before using the randomization function the list is filled using all available examples
                If Grapheme.ExamplesWithPhonemeReplacement.Count < MaximumNumberOfExampleWords Then
                    Grapheme.ExamplesWithPhonemeReplacement.Add(OrthographicForm & PhoneticFormConcat & " " & OutputTranscription)
                Else
                    'Testing if RandomizedPhonemeIndex is within the index range of the current phoneme combination
                    If RandomizedPhonemeIndex >= PhoneIndex And RandomizedPhonemeIndex < PhoneIndex + PhonemeLength Then
                        'Or else continues filling the list using the randomization function
                        If Grapheme.ExampleSamplingWithPhonemeReplacementActive = True Then
                            Grapheme.ExamplesWithPhonemeReplacement.Add(OrthographicForm & PhoneticFormConcat & " " & OutputTranscription)
                            Grapheme.ExampleSamplingWithPhonemeReplacementActive = False
                            Grapheme.ExamplesWithPhonemeReplacement.RemoveAt(1)
                        End If
                    End If
                End If

                'Marking the word as containing PDS addition
                ManualEvaluations.Add("Phoneme replacement of " & OriginalPhoneme & " -> " & CurrentPhonemeString)

                'Also exporting these to log file for evaluation
                SendInfoToLog(OrthographicForm & vbTab & OutputTranscription & vbTab &
                              OutputSpelling & PhoneticFormConcat & vbTab &
                              OriginalPhoneme & vbTab & CurrentPhonemeString, "PhonemeReplacements", OutputFolder)

                'End If
            End If

            'Adding Silent grapheme exmaples
            If PhonemeLength = 0 Then
                'before using the randomization function the list is filled using all available examples
                If Grapheme.ExamplesOfSilentgraphemes.Count < MaximumNumberOfExampleWords Then
                    Grapheme.ExamplesOfSilentgraphemes.Add(OrthographicForm & PhoneticFormConcat & " " & OutputTranscription)
                Else
                    'Testing if RandomizedPhonemeIndex is within the index range of the current phoneme combination
                    If RandomizedPhonemeIndex >= PhoneIndex And RandomizedPhonemeIndex < PhoneIndex + PhonemeLength Then
                        'Or else continues filling the list using the randomization function
                        If Grapheme.ExampleSamplingWithSilentgraphemesActive = True Then
                            Grapheme.ExamplesOfSilentgraphemes.Add(OrthographicForm & PhoneticFormConcat & " " & OutputTranscription)
                            Grapheme.ExampleSamplingWithSilentgraphemesActive = False
                            Grapheme.ExamplesOfSilentgraphemes.RemoveAt(1)
                        End If
                    End If
                End If

                'Marking the word as containing PDS addition
                ManualEvaluations.Add("Silent grapheme: " & Grapheme.PossibleSpelling)

                'Also exporting these to log file for evaluation
                SendInfoToLog(OrthographicForm & vbTab & OutputTranscription & vbTab &
                              OutputSpelling & PhoneticFormConcat & vbTab &
                              OriginalPhoneme & vbTab & CurrentPhonemeString, "SilentGrapheme_" & Grapheme.PossibleSpelling, OutputFolder)

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub

    Private Sub p2g_MarkAndExportMostCommonExampleWords(ByRef Tempspelling As String,
                                                  ByRef PhoneticForm As List(Of String),
                                                  ByRef PhoneToGraphemesDictionary As PhoneToGraphemes,
                                                  ByRef PDCS_Grapheme As Boolean,
                                                  ByRef PDS_Grapheme As Boolean,
                                                  ByRef PR_Grapheme As Integer,
                                                  ByRef OriginalPhoneme As String,
                                                  ByRef UseReducedPhoneticForm As Boolean,
                                                            ByRef CurrentPhonemeString As String,
                                                  ByRef Grapheme As PhoneToGraphemes.Grapheme,
                                                  ByRef PhonemeLength As Integer,
                                                  ByRef OutputFolder As String,
                                                           ByRef PhonemeHistory As List(Of String))

        Dim UsePhonotacticProbabilityCriterion As Boolean = True 'If set to true the word with the highest average phonotactic probability will be used as example
        'If set to false, the word with the highest frequency of occurence will be used instead

        Try

            'Stores the most common example word, that use no special function, or only one special functions. Word that use more than one special function never stored as examples.
            'Adding normal example words (that do not use any special function, or silent graphemes)
            Dim WordContainsSilentgrapheme As Boolean = False
            If PhonemeHistory.Contains(ZeroPhoneme) Then
                WordContainsSilentgrapheme = True
            End If

            If PDCS_Grapheme = False And PDS_Grapheme = False And PR_Grapheme = False And WordContainsSilentgrapheme = False Then
                'Adds the word only if it fulfils the conditions in AddExampleWord
                AddExampleWord(Grapheme.ExampleWordWithoutPDCS)
            End If

            'Adding PDCS examples
            If PDCS_Grapheme = True And PDS_Grapheme = False And PR_Grapheme = False And WordContainsSilentgrapheme = False Then
                'Adds the word only if it fulfils the conditions in AddExampleWord
                AddExampleWord(Grapheme.ExampleWordWithPDCS)
            End If

            'Adding PDS examples
            If PDCS_Grapheme = False And PDS_Grapheme = True And PR_Grapheme = False And WordContainsSilentgrapheme = False Then
                'Adds the word only if it fulfils the conditions in AddExampleWord
                AddExampleWord(Grapheme.ExampleWordWithPDS)
            End If

            'Adding Phoneme replacement examples
            If PDCS_Grapheme = False And PDS_Grapheme = False And PR_Grapheme = True And WordContainsSilentgrapheme = False Then
                'Adds the word only if it fulfils the conditions in AddExampleWord
                AddExampleWord(Grapheme.ExampleWordWithPhonemeReplacement)
            End If

            'Adding Silent grapheme exmaples
            If PDCS_Grapheme = False And PDS_Grapheme = False And PR_Grapheme = False And PhonemeLength = 0 Then
                'Adds the word only if it fulfils the conditions in AddExampleWord
                AddExampleWord(Grapheme.ExampleWordOfSilentgraphemes)
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub


    ''' <summary>
    ''' Determines if an example word should be added i p2g_CheckNext, if so, the sub, adds the word.
    ''' </summary>
    Private Sub AddExampleWord(ByRef ExampleWord As Word)

        'Dim ExampleWordFrequencyLimit As Integer = 10

        'Not adding Hapax legomenas
        'If Not RawWordTypeFrequency >= ExampleWordFrequencyLimit Then 'This limit is skipped

        'Not adding example words if they contain an unresolved character (as this may acctually cause erroneous segmentation)
        'If Not String.Concat(Graphemes).Contains(Unresolved_p2g_Character) Then
        'This is already evaluated in the calling code

        'Adding the word if no word has been previously added
        If ExampleWord Is Nothing Then
            ExampleWord = Me
        Else

            'Replacing the existing word if it is marked as foreign, and the new one is not
            If ExampleWord.ForeignWord = True And ForeignWord = False Then
                ExampleWord = Me
            Else

                'Replacing the existing word if contains a special character, and new one is not
                If ExampleWord.OrthographicFormContainsSpecialCharacter = True And OrthographicFormContainsSpecialCharacter = False Then
                    ExampleWord = Me
                Else

                    'Replaces the existing word if it has a phoneme range is 3-8 phonemes and a higher phonotactic probability (for publication purpose)
                    Dim PhonemeCount As Integer = Me.CountPhonemes
                    If PhonemeCount > 2 And PhonemeCount < 9 Then

                        Dim NewWordValues As New List(Of Double) From {SSPP_Average, RawWordTypeFrequency}
                        Dim ExistingWordValues As New List(Of Double) From {ExampleWord.SSPP_Average,
                                                ExampleWord.RawWordTypeFrequency}

                        If GeometricMean(NewWordValues) > GeometricMean(ExistingWordValues) Then
                            ExampleWord = Me
                        End If
                    End If
                End If
            End If
        End If

    End Sub


#End Region

#Region "Neighborhood"


    <Serializable>
    Public Class LD_String
        Implements IComparable

        Public Sub New()

        End Sub
        'Implements IComparable(Of String)

        Public Property LevenshteinDifference As Integer
        Public Property Word As New Word
        'Public Property OrthographicForm As String
        'Public Property Phonemes As New List(Of String)

        'Public Function CompareTo(other As String) As Integer Implements IComparable(Of String).CompareTo
        'Return OrthographicForm.CompareTo(other)
        'End Function

        Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            If Not TypeOf (obj) Is LD_String Then
                Throw New ArgumentException()
            Else

                Dim tempOLD As LD_String = DirectCast(obj, LD_String)

                If Me.LevenshteinDifference < tempOLD.LevenshteinDifference Then
                    Return -1
                ElseIf Me.LevenshteinDifference = tempOLD.LevenshteinDifference Then
                    Return 0
                Else
                    Return 1
                End If
            End If

        End Function

    End Class

    <Serializable>
    Public Class PLD
        Implements IComparable(Of List(Of String))
        Implements IComparable

        Public Property LevenshteinDifference As Short
        Public Property Word As New Word
        'Public Property OrthographicForm As String
        'Public Property Phonemes As New List(Of String)

        Public Function CompareTo(other As List(Of String)) As Integer Implements IComparable(Of List(Of String)).CompareTo
            Throw New NotImplementedException()
        End Function


        Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            If Not TypeOf (obj) Is PLD Then
                Throw New ArgumentException()
            Else
                Dim tempPLD As PLD = DirectCast(obj, PLD)

                If Me.LevenshteinDifference < tempPLD.LevenshteinDifference Then
                    Return -1
                ElseIf Me.LevenshteinDifference = tempPLD.LevenshteinDifference Then
                    Return 0
                Else
                    Return 1
                End If
            End If
        End Function

    End Class




    Public Enum LevenShteinDistanceTypes
        OLD
        PLD_Full
        PLD_Reduced
    End Enum

    ''' <summary>
    ''' This sub calculates the Levenshtein distance to the strings in a wordlist determined by the type. All words with a LD of 1 are 
    ''' stored in the property LDWords, as well as all words up to MinimumWordCount irrespective of their LD.
    ''' The sub also automatically calculates and sets OLD1/PLD1_Count, FrequencyWeightedDensity_OLD/PLD, OLD/PLD20_Mean (depending on type).
    ''' </summary>
    ''' <param name="type">Determines id orthographic (OLD) or phonetic (PLD) Levenshtein distance should be calculated.</param>
    ''' <param name="MinimumWordCount">The number of words to be included when calculating mean OLD of the MinimumWordCount closest neighbours (E.G. use 20 to calculate OLD20).</param>
    Public Sub LevenshteinDistance(ByVal type As LevenShteinDistanceTypes, Optional ByVal MinimumWordCount As Integer = 20)

        If OLD_Data Is Nothing Then OLD_Data = New OrthLD_Data_Type
        If PLD_Data Is Nothing Then PLD_Data = New PLD_Data_Type

        Dim comparisonCorpus As String() = {}

        Select Case type
            Case LevenShteinDistanceTypes.OLD
                comparisonCorpus = OLD_corpus

            Case LevenShteinDistanceTypes.PLD_Full
                comparisonCorpus = PLD_IPA_Corpus

            Case Else
                Throw New NotImplementedException

        End Select

        Dim comparisonCorpusLength As Integer = comparisonCorpus.Length

        Try


            Dim LDs As New List(Of WordGroup.LD_String)

            Dim startTime As DateTime = DateTime.Now

            Dim firstPartLength As Integer = MinimumWordCount
            If comparisonCorpusLength < MinimumWordCount Then firstPartLength = comparisonCorpusLength

            For i = 0 To firstPartLength - 1
                Dim NewLD As New WordGroup.LD_String
                Dim CurrentComparisonCorpusForm As String = comparisonCorpus(i).Split(vbTab)(0)

                Select Case type
                    Case LevenShteinDistanceTypes.OLD
                        NewLD.LevenshteinDifference = MathMethods.LevenshteinDistance(OrthographicForm, CurrentComparisonCorpusForm)
                        NewLD.Word.OrthographicForm = CurrentComparisonCorpusForm
                    Case LevenShteinDistanceTypes.PLD_Full
                        NewLD.LevenshteinDifference = MathMethods.LevenshteinDistance(TranscriptionString, CurrentComparisonCorpusForm)
                        NewLD.Word.TranscriptionString = CurrentComparisonCorpusForm
                End Select

                LDs.Add(NewLD)

            Next

            'Sorts the first MinimumWordCount OLDs
            LDs.Sort()

            'Goes through the rest of the comparisonCorpus
            For i = firstPartLength To comparisonCorpusLength - 1
                Dim CurrentComparisonCorpusForm As String = comparisonCorpus(i).Split(vbTab)(0)
                Dim LD As Short
                Select Case type
                    Case LevenShteinDistanceTypes.OLD
                        LD = MathMethods.LevenshteinDistance(OrthographicForm, CurrentComparisonCorpusForm)
                    Case LevenShteinDistanceTypes.PLD_Full
                        LD = MathMethods.LevenshteinDistance(TranscriptionString, CurrentComparisonCorpusForm)
                End Select


                If Not LD = 0 Then 'I.E. it'd be the same wordtype
                    If LD = 1 Or LD < LDs(MinimumWordCount - 1).LevenshteinDifference Then

                        'Adding a new LD word if is is a LD1 word, or if it has a lower LD than the word on place MinimumWordCount-1
                        Dim NewLD As New WordGroup.LD_String

                        Select Case type
                            Case LevenShteinDistanceTypes.OLD
                                NewLD.Word.OrthographicForm = CurrentComparisonCorpusForm
                            Case LevenShteinDistanceTypes.PLD_Full
                                NewLD.Word.TranscriptionString = CurrentComparisonCorpusForm
                        End Select

                        NewLD.LevenshteinDifference = LD
                        LDs.Add(NewLD)
                        LDs.Sort()

                        'Removes the last LD word if it is outside the desired minimum LD words count and is higher than LD=1. Storing the LD1 words in the list
                        If LDs.Count > MinimumWordCount And LDs(LDs.Count - 1).LevenshteinDifference > 1 Then
                            LDs.RemoveAt(LDs.Count - 1)
                        End If
                    End If
                End If
            Next

            'Storing data
            Select Case type
                Case LevenShteinDistanceTypes.OLD
                    OLD_Data.OLD_Words = LDs

                    'Counting the number of LD1 words, and calculating frequency weighted density
                    Dim count As Integer = 0
                    Dim freqWeightedDensitySum As Double = 0
                    For n = 0 To OLD_Data.OLD_Words.Count - 1
                        If OLD_Data.OLD_Words(n).LevenshteinDifference = 1 Then
                            count += 1
                            freqWeightedDensitySum += OLD_Data.OLD_Words(n).Word.ZipfValue_Word
                        End If
                    Next
                    OLD_Data.OLD1_Count = count
                    FrequencyWeightedDensity_OLD = freqWeightedDensitySum

                    'Calculating average PLD for the MinimumWordCount closest neighbours
                    Dim sum As Integer = 0
                    For n = 0 To MinimumWordCount - 1
                        sum += LDs(n).LevenshteinDifference
                    Next
                    OLD_Data.OLD20_Mean = sum / MinimumWordCount

                Case LevenShteinDistanceTypes.PLD_Full
                    PLD_Data.PLD_Words = LDs

                    'Counting the number of LD1 words, and calculating frequency weighted density
                    Dim count As Integer = 0
                    Dim freqWeightedDensitySum As Double = 0
                    For n = 0 To PLD_Data.PLD_Words.Count - 1
                        If PLD_Data.PLD_Words(n).LevenshteinDifference = 1 Then
                            count += 1
                            freqWeightedDensitySum += PLD_Data.PLD_Words(n).Word.ZipfValue_Word
                        End If
                    Next
                    PLD_Data.PLD1_Count = count
                    FWPN_DensityProbability = freqWeightedDensitySum

                    'Calculating average PLD for the MinimumWordCount closest neighbours
                    Dim sum As Integer = 0
                    For n = 0 To MinimumWordCount - 1
                        sum += LDs(n).LevenshteinDifference
                    Next
                    PLD_Data.PLD20_Mean = sum / MinimumWordCount

            End Select

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub


#End Region

#Region "Sorting"


    ''' <summary>
    ''' Sorts the items in AllPossiblePoS after word frequency, in falling order.
    ''' </summary>
    Public Sub SortAllPossiblePoS()

        Dim Query1 = AllPossiblePoS.OrderByDescending(Function(PoSItem) PoSItem.Item2)
        Dim SortedList As New List(Of Tuple(Of String, Double))

        'Adding in sorted order
        For Each PoSItem In Query1
            SortedList.Add(PoSItem)
        Next

        'Clearing AllPossiblePoS
        AllPossiblePoS.Clear()

        'Putting sorted items back
        For n = 0 To SortedList.Count - 1
            AllPossiblePoS.Add(SortedList(n))
        Next

    End Sub

    ''' <summary>
    ''' Sorts the items in AllPossiblePoS after word frequency, in falling order.
    ''' </summary>
    Public Sub SortAllOccurringLemmas()

        Dim Query1 = AllOccurringLemmas.OrderByDescending(Function(LemmaItem) LemmaItem.Item2)
        Dim SortedList As New List(Of Tuple(Of String, Double))

        'Adding in sorted order
        For Each LemmaItem In Query1
            SortedList.Add(LemmaItem)
        Next

        'Clearing AllPossiblePoS
        AllOccurringLemmas.Clear()

        'Putting sorted items back
        For n = 0 To SortedList.Count - 1
            AllOccurringLemmas.Add(SortedList(n))
        Next

    End Sub


    ''' <summary>
    ''' Sorts the items in PLD1Transcriptions after word frequency (Zipf-scale value), in falling order.
    ''' </summary>
    ''' <param name="LeaveFirstItemAtIndexZero">If set to false the whole PLD1 list is sorted. If left to true, the first item is assumed to be the PLD1 transcriptino of the current word, which is hence left at index zero.</param>
    Public Sub SortPLD1Transcriptions(Optional ByVal LeaveFirstItemAtIndexZero As Boolean = True)

        'Exits straight away if there are less than two items to sort
        If PLD1Transcriptions.Count < 2 Then Exit Sub

        'Stores the first item and temporaliy removes it from the list, in order to sort the list without it, and then reinserting it again below at index 0.
        Dim FirstItem As String = ""
        If LeaveFirstItemAtIndexZero = True Then
            FirstItem = PLD1Transcriptions(0)
            PLD1Transcriptions.RemoveAt(0)
        End If

        'Sorting the remainder of the PLD1 list
        'Dim ListToSort As New List(Of StringDoubleCombination)
        'For n = 0 To PLD1Transcriptions.Count - 1
        '    Dim PLD1TranscriptionSplit() As String = PLD1Transcriptions(n).Split(":")
        '    Dim NewPLD1Transcription As New StringDoubleCombination With {.StringData = PLD1TranscriptionSplit(0)}
        '    If PLD1TranscriptionSplit.Length > 0 Then NewPLD1Transcription.DoubleData = PLD1TranscriptionSplit(1)
        '    ListToSort.Add(NewPLD1Transcription)
        'Next

        Dim ListToSort As New List(Of Tuple(Of String, Double))
        For n = 0 To PLD1Transcriptions.Count - 1
            Dim PLD1TranscriptionSplit() As String = PLD1Transcriptions(n).Split(":")
            If PLD1TranscriptionSplit.Length = 1 Then
                ListToSort.Add(New Tuple(Of String, Double)(PLD1TranscriptionSplit(0), 0))
            ElseIf PLD1TranscriptionSplit.Length > 1 Then
                ListToSort.Add(New Tuple(Of String, Double)(PLD1TranscriptionSplit(0), PLD1TranscriptionSplit(1)))
            End If
        Next

        Dim Query1 = ListToSort.OrderByDescending(Function(PLD1TranscriptionItem) PLD1TranscriptionItem.Item2)
        Dim SortedList As New List(Of Tuple(Of String, Double))

        'Adding in sorted order
        For Each PLD1TranscriptionItem In Query1
            SortedList.Add(PLD1TranscriptionItem)
        Next

        'Clearing PLD1Transcriptions
        PLD1Transcriptions.Clear()

        'Putting sorted items back
        For n = 0 To SortedList.Count - 1
            PLD1Transcriptions.Add(SortedList(n).Item1 & ":" & SortedList(n).Item2)
        Next

        'Reinserting first item at index 0
        If LeaveFirstItemAtIndexZero = True Then
            PLD1Transcriptions.Insert(0, FirstItem)
        End If

    End Sub


    ''' <summary>
    ''' Sorts the items in OLD1Spellings after word frequency (Zipf-scale value), in falling order.
    ''' </summary>
    ''' <param name="LeaveFirstItemAtIndexZero">If set to false the whole OLD1 list is sorted. If left to true, the first item is assumed to be the spelling of the current word, which is hence left at index zero.</param>
    Public Sub SortOLD1Spellings(Optional ByVal LeaveFirstItemAtIndexZero As Boolean = True)

        'Exits straight away if there are less than two items to sort
        If OLD1Spellings.Count < 2 Then Exit Sub

        'Stores the first item and temporaliy removes it from the list, in order to sort the list without it, and then reinserting it again below at index 0.
        Dim FirstItem As String = ""
        If LeaveFirstItemAtIndexZero = True Then
            FirstItem = OLD1Spellings(0)
            OLD1Spellings.RemoveAt(0)
        End If

        'Sorting the remainder of the OLD1 list
        'Dim ListToSort As New List(Of StringDoubleCombination)
        'For n = 0 To OLD1Spellings.Count - 1
        '    Dim OLD1SpellingSplit() As String = OLD1Spellings(n).Split(",") 'Comma is used instead of colon since is a part of some spellings.
        '    Dim NewOLD1Spelling As New StringDoubleCombination With {.StringData = OLD1SpellingSplit(0)}
        '    If OLD1SpellingSplit.Length > 0 Then NewOLD1Spelling.DoubleData = OLD1SpellingSplit(1)
        '    ListToSort.Add(NewOLD1Spelling)
        'Next

        Dim ListToSort As New List(Of Tuple(Of String, Double))
        For n = 0 To OLD1Spellings.Count - 1
            Dim OLD1SpellingSplit() As String = OLD1Spellings(n).Split(",") 'Comma is used instead of colon since is a part of some spellings.
            If OLD1SpellingSplit.Length = 1 Then
                ListToSort.Add(New Tuple(Of String, Double)(OLD1SpellingSplit(0), 0))
            ElseIf OLD1SpellingSplit.Length > 1 Then
                ListToSort.Add(New Tuple(Of String, Double)(OLD1SpellingSplit(0), OLD1SpellingSplit(1)))
            End If
        Next


        Dim Query1 = ListToSort.OrderByDescending(Function(OLD1SpellingItem) OLD1SpellingItem.Item2)
        Dim SortedList As New List(Of Tuple(Of String, Double))

        'Adding in sorted order
        For Each OLD1SpellingItemItem In Query1
            SortedList.Add(OLD1SpellingItemItem)
        Next

        'Clearing OLD1Spellings
        OLD1Spellings.Clear()

        'Putting sorted items back
        For n = 0 To SortedList.Count - 1
            OLD1Spellings.Add(SortedList(n).Item1 & "," & SortedList(n).Item2)
        Next

        'Reinserting first item at index 0
        If LeaveFirstItemAtIndexZero = True Then
            OLD1Spellings.Insert(0, FirstItem)
        End If

    End Sub


#End Region

#Region "InputOutput"


    ''' <summary>
    ''' Returns a new Word which is a deep copy of the original.
    ''' </summary>
    ''' <returns>Returns a new Word which is a deep copy of the original.</returns>
    Public Function CreateCopy() As Word

        'Creating an output object
        Dim newWord As Word

        'Serializing to memorystream
        Dim serializedMe As New MemoryStream
        Dim serializer As New BinaryFormatter
        serializer.Serialize(serializedMe, Me)

        'Deserializing to new object
        serializedMe.Position = 0
        newWord = CType(serializer.Deserialize(serializedMe), Word)
        serializedMe.Close()

        'Returning the new object
        Return newWord


    End Function

    ''' <summary>
    ''' Returns a new Word which is a shallow copy of the original.
    ''' </summary>
    ''' <returns>Returns a new Word which is a shallow copy of the original.</returns>
    Public Function CreateShallowCopy() As Word

        'Creating an output object
        Dim newWord As Word

        'Creating a shallow copy
        newWord = MemberwiseClone()

        'Returning the new object
        Return newWord

    End Function

    ''' <summary>Returns a New Word which Is a shallow copy Of the original. Only value type And String objects are copied. 
    ''' No other references to any other reference type objects are copied.</summary>
    ''' <returns>Returns a new Word which is a shallow copy of the original. Only value type and string objects are copied. 
    ''' No other references to any other reference type objects are copied.</returns>
    Public Function CreateSuperShallowCopy() As Word

        'Creating an output object
        Dim newWord As New Word

        'Creating a Super shallow copy

        'Getting a list of properties of the Word class
        Dim pinfo() As PropertyInfo = GetType(Word).GetProperties()
        For p = 0 To pinfo.Length - 1

            'Copies any string properties
            If pinfo(p).PropertyType = GetType(String) Then
                pinfo(p).SetValue(newWord, pinfo(p).GetValue(Me))
            End If

            'Copies all other non reference type properties
            If IsReference(pinfo(p).GetValue(Me)) = False Then
                pinfo(p).SetValue(newWord, pinfo(p).GetValue(Me))
            End If

        Next

        'Returning the new object
        Return newWord

    End Function




    ''' <summary>
    ''' Generates a string repressenting the orthigraphic and phonetic info on a word, which can be parsed with the function ParseInputWordString.
    ''' </summary>
    ''' <param name="AllowIrrelevantValues">If set to false, fields that contain irrelevant data, (such as P2G data when no phonetic transcription exist) are exported as empty fields. If left to false, a default or irrelevant value is exported for some data types.</param>
    '''<param name="GenerateLackingData">If set to true, some data is automatically generated if lacking.</param>
    ''' <returns></returns>
    Public Function GenerateFullPhoneticOutputTxtString(ByRef ColumnOrder As PhoneticTxtStringColumnIndices,
                                                        Optional ByVal RoundingDecimals As Integer = 4,
                                                        Optional ByVal AllowIrrelevantValues As Boolean = True,
                                                        Optional ByVal GenerateLackingData As Boolean = False) As String

        'Counting phones to detemine if a phonetic transcription exists. This value is used to skip fields that are based on the existance of a phonetic transcription, such as sonographs
        Dim LocalPhoneCount As Integer = CountPhonemes(False)


        'Creating phonemeArrays, only if LocalPhoneCount > 0 
        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then GeneratePhoneticForms()

        Dim outputString As String = ""
        Dim ColumnToWrite As Integer = 0

        'Counting the number of columns to export
        Dim ColumnOrderProperyInfo() As System.Reflection.PropertyInfo = GetType(PhoneticTxtStringColumnIndices).GetProperties
        For n = 0 To ColumnOrderProperyInfo.Length - 1
            If GetType(PhoneticTxtStringColumnIndices).GetProperty(ColumnOrderProperyInfo(n).Name).GetValue(ColumnOrder) IsNot Nothing Then
                If GetType(PhoneticTxtStringColumnIndices).GetProperty(ColumnOrderProperyInfo(n).Name).GetValue(ColumnOrder) = ColumnToWrite Then
                    outputString &= ""

                    If ColumnOrderProperyInfo(n).Name = "OrthographicForm" Then outputString &= OrthographicForm & vbTab

                    If ColumnOrderProperyInfo(n).Name = "GIL2P_OT_Average" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Rounding(GIL2P_OT_Average, , RoundingDecimals) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "GIL2P_OT_Min" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Rounding(GIL2P_OT_Min, , RoundingDecimals) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PIP2G_OT_Average" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Rounding(PIP2G_OT_Average, , RoundingDecimals) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PIP2G_OT_Min" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Rounding(PIP2G_OT_Min, , RoundingDecimals) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "G2P_OT_Average" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Rounding(G2P_OT_Average,, RoundingDecimals) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "UpperCase" Then
                        If ContainsWordListData = True Then
                            outputString &= Rounding(ProportionStartingWithUpperCase,, 2) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "Homographs" Then
                        If LanguageHomographs IsNot Nothing Then
                            outputString &= String.Join("|", LanguageHomographs) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If
                    If ColumnOrderProperyInfo(n).Name = "HomographCount" Then outputString &= LanguageHomographCount & vbTab

                    If ColumnOrderProperyInfo(n).Name = "SpecialCharacter" Then
                        If ContainsWordListData = True Then
                            outputString &= OrthographicFormContainsSpecialCharacter & vbTab
                        Else
                            If GenerateLackingData = True Then
                                outputString &= MarkSpecialCharacterWords_ByNormalCharList(SwedishOrthographicCharacters) & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "RawWordTypeFrequency" Then
                        If ContainsWordListData = True Then
                            outputString &= RawWordTypeFrequency & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "RawDocumentCount" Then
                        If ContainsWordListData = True Then
                            outputString &= RawDocumentCount & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PhoneticForm" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= String.Join(" ", BuildExtendedIpaArray(False)) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    'ReadOnly line (any data here is exported even if no ordinary phonetic transcription exists)
                    If ColumnOrderProperyInfo(n).Name = "TemporarySyllabification" Then outputString &= String.Join(" ", BuildExtendedIpaArray(False,,,, True)) & vbTab

                    If ColumnOrderProperyInfo(n).Name = "ReducedTranscription" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= String.Join(" ", BuildReducedIpaArray()) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PhonotacticType" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            'Checks if phonotactic type has been determined
                            If PhonotacticType <> "" Then
                                outputString &= PhonotacticType & vbTab
                            Else
                                If GenerateLackingData = True Then
                                    'Determines the phonotectic type
                                    outputString &= SetWordPhonotacticType() & vbTab
                                Else
                                    outputString &= vbTab
                                End If
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "SSPP_Average" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Rounding(SSPP_Average, , RoundingDecimals) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "SSPP_Min" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Rounding(SSPP_Min,, RoundingDecimals) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PSP_Sum" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Rounding(PSP_Sum,, RoundingDecimals) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PSBP_Sum" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Rounding(PSBP_Sum,, RoundingDecimals) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "S_PSP_Average" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Rounding(S_PSP_Average,, RoundingDecimals) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "S_PSBP_Average" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Rounding(S_PSBP_Average,, RoundingDecimals) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "Homophones" Then
                        If LanguageHomophones IsNot Nothing Then
                            outputString &= String.Join("|", LanguageHomophones) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "HomophoneCount" Then outputString &= LanguageHomophoneCount & vbTab

                    If ColumnOrderProperyInfo(n).Name = "PNDP" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Rounding(FWPN_DensityProbability, , 3) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PLD1Transcriptions" Then
                        If PLD1Transcriptions.Count > 0 Then
                            outputString &= String.Join("|", PLD1Transcriptions) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "ONDP" Then
                        'If AllowIrrelevantValues = True Then
                        outputString &= Rounding(FWON_DensityProbability, , 3) & vbTab
                        'Else
                        'outputString &= vbTab
                        'End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "OLD1Spellings" Then

                        'Creating OLD1Spellings since they may not exist in all serialized versions...
                        If OLD1Spellings Is Nothing Then OLD1Spellings = New List(Of String)

                        If OLD1Spellings.Count > 0 Then
                            outputString &= String.Join("|", OLD1Spellings) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PLDx_Average" Then
                        'Exporting PLDx_Average only if there exists PLDx words
                        If PLDxData.Count > 0 Then
                            outputString &= Rounding(PLDx_Average, , RoundingDecimals) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "OLDx_Average" Then
                        If OLDxData.Count > 0 Then
                            outputString &= Rounding(OLDx_Average, , RoundingDecimals) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PLDx_Neighbors" Then
                        'Creating PLDxData since it may not exist in all serialized versions...
                        If PLDxData Is Nothing Then PLDxData = New List(Of Tuple(Of Integer, String, Single))

                        If PLDxData.Count > 0 Then
                            Dim PLDxOutputList As New List(Of String)
                            For i = 0 To PLDxData.Count - 1
                                PLDxOutputList.Add(PLDxData(i).Item1 & ":" & PLDxData(i).Item2 & ":" & PLDxData(i).Item3)
                            Next
                            outputString &= String.Join("|", PLDxOutputList) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "OLDx_Neighbors" Then
                        'Creating OLDxData since it may not exist in all serialized versions...
                        If OLDxData Is Nothing Then OLDxData = New List(Of Tuple(Of Integer, String, Single))

                        If OLDxData.Count > 0 Then
                            Dim OLDxOutputList As New List(Of String)
                            For i = 0 To OLDxData.Count - 1
                                OLDxOutputList.Add(OLDxData(i).Item1 & ":" & OLDxData(i).Item2 & ":" & OLDxData(i).Item3)
                            Next
                            outputString &= String.Join("|", OLDxOutputList) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "Sonographs" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then

                            Dim LocalSonographs As New List(Of String)
                            For i = 0 To Sonographs_Letters.Count - 1
                                LocalSonographs.Add(Sonographs_Letters(i) & "-" & Sonographs_Pronunciation(i))
                            Next

                            If LocalSonographs.Count > 0 Then
                                outputString &= String.Join("|", LocalSonographs) & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "AllPoS" Then
                        If ContainsWordListData = True Then
                            'Preparing PoS string
                            Dim PartsOfSpeechString As String = ""
                            For PoS = 0 To AllPossiblePoS.Count - 1
                                PartsOfSpeechString &= AllPossiblePoS(PoS).Item1 & ":" & Rounding(AllPossiblePoS(PoS).Item2, , 2) & "|"
                            Next
                            'Removes the last comma
                            PartsOfSpeechString = PartsOfSpeechString.TrimEnd("|")

                            'Writing PoS string
                            outputString &= PartsOfSpeechString & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "AllLemmas" Then
                        If ContainsWordListData = True Then
                            'Preparing lemma string
                            Dim LemmaString As String = ""
                            For PoS = 0 To AllOccurringLemmas.Count - 1
                                LemmaString &= AllOccurringLemmas(PoS).Item1 & ":" & Rounding(AllOccurringLemmas(PoS).Item2, , 2) & "|"
                            Next
                            'Removes the last comma
                            LemmaString = LemmaString.TrimEnd("|")
                            'Writing lemma string
                            outputString &= LemmaString & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "NumberOfSenses" Then
                        If ContainsWordListData = True Then
                            If NumberOfSenses > 0 Then
                                outputString &= NumberOfSenses & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "Abbreviation" Then
                        If ContainsWordListData = True Then
                            outputString &= Abbreviation & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "Acronym" Then
                        If ContainsWordListData = True Then
                            outputString &= Acronym & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "AllPossibleSenses" Then
                        If AllPossibleSenses.Count > 0 Then
                            outputString &= String.Join("|", AllPossibleSenses) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "SSPP" Then

                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then

                            If PP IsNot Nothing And PP_Phonemes IsNot Nothing Then

                                Dim PhonoTacticCombination As New List(Of String)
                                Dim i As Integer = 0
                                For x = 0 To PP_Phonemes.Count - 1
                                    Dim probString As String = ""
                                    If Not i > PP_Phonemes.Count - 1 Then
                                        probString = "[" & PP_Phonemes(i) & "] "
                                    End If
                                    If Not i > PP.Count - 1 Then
                                        probString &= Rounding(PP(i),, 4)
                                    End If
                                    PhonoTacticCombination.Add(probString)
                                    i += 1
                                Next

                                'Adding also a word end string
                                If Not i > PP_Phonemes.Count - 1 Then
                                    PhonoTacticCombination.Add("[" & PP_Phonemes(i) & "]")
                                End If

                                outputString &= String.Join(" ", PhonoTacticCombination) & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PSP" Then

                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then

                            If PSP IsNot Nothing And PSP_Phonemes IsNot Nothing Then

                                Dim PhonoTacticCombination As New List(Of String)
                                Dim i As Integer = 0
                                For x = 0 To PSP_Phonemes.Count - 1
                                    Dim probString As String = ""
                                    If Not i > PSP_Phonemes.Count - 1 Then
                                        probString = "[" & PSP_Phonemes(i) & "] "
                                    End If
                                    If Not i > PSP.Count - 1 Then
                                        probString &= Rounding(PSP(i),, 4)
                                    End If
                                    PhonoTacticCombination.Add(probString)
                                    i += 1
                                Next

                                outputString &= String.Join(" ", PhonoTacticCombination) & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PSBP" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            If PSBP IsNot Nothing And PSBP_Phonemes IsNot Nothing Then

                                Dim PhonoTacticCombination As New List(Of String)
                                Dim i As Integer = 0
                                For x = 0 To PSBP_Phonemes.Count - 1
                                    Dim probString As String = ""
                                    If Not i > PSBP_Phonemes.Count - 1 Then
                                        probString = "[" & PSBP_Phonemes(i) & "] "
                                    End If
                                    If Not i > PSBP.Count - 1 Then
                                        probString &= Rounding(PSBP(i),, 6)
                                    End If
                                    PhonoTacticCombination.Add(probString)
                                    i += 1
                                Next

                                outputString &= String.Join(" ", PhonoTacticCombination) & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "S_PSP" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            If S_PSP IsNot Nothing And S_PSP_Phonemes IsNot Nothing Then

                                Dim PhonoTacticCombination As New List(Of String)
                                Dim i As Integer = 0
                                For x = 0 To S_PSP_Phonemes.Count - 1
                                    Dim probString As String = ""
                                    If Not i > S_PSP_Phonemes.Count - 1 Then
                                        probString = "[" & S_PSP_Phonemes(i) & "] "
                                    End If
                                    If Not i > S_PSP.Count - 1 Then
                                        probString &= Rounding(S_PSP(i),, 4)
                                    End If
                                    PhonoTacticCombination.Add(probString)
                                    i += 1
                                Next

                                outputString &= String.Join(" ", PhonoTacticCombination) & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "S_PSBP" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            If S_PSBP IsNot Nothing And S_PSBP_Phonemes IsNot Nothing Then

                                Dim PhonoTacticCombination As New List(Of String)
                                Dim i As Integer = 0
                                For x = 0 To S_PSBP_Phonemes.Count - 1
                                    Dim probString As String = ""
                                    If Not i > S_PSBP_Phonemes.Count - 1 Then
                                        probString = "[" & S_PSBP_Phonemes(i) & "] "
                                    End If
                                    If Not i > S_PSBP.Count - 1 Then
                                        probString &= Rounding(S_PSBP(i),, 6)
                                    End If
                                    PhonoTacticCombination.Add(probString)
                                    i += 1
                                Next

                                outputString &= String.Join(" ", PhonoTacticCombination) & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If


                    If ColumnOrderProperyInfo(n).Name = "GIL2P_OT" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            Dim GIL2P_OT_ExportValues As New List(Of Double)
                            'Rounding the output values
                            For r = 0 To GIL2P_OT.Count - 1
                                GIL2P_OT_ExportValues.Add(Rounding(GIL2P_OT(r),, RoundingDecimals))
                            Next

                            If GIL2P_OT_ExportValues.Count > 0 Then
                                outputString &= String.Join(" ", GIL2P_OT_ExportValues) & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PIP2G_OT" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            Dim PIP2G_OT_ExportValues As New List(Of Double)
                            'Rounding the output values
                            For r = 0 To PIP2G_OT.Count - 1
                                PIP2G_OT_ExportValues.Add(Rounding(PIP2G_OT(r),, RoundingDecimals))
                            Next

                            If PIP2G_OT_ExportValues.Count > 0 Then
                                outputString &= String.Join(" ", PIP2G_OT_ExportValues) & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "G2P_OT" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            Dim SpellingProbability_g2p_ExportValues As New List(Of Double)
                            'Rounding the spelling regularity values
                            For r = 0 To G2P_OT.Count - 1
                                SpellingProbability_g2p_ExportValues.Add(Rounding(G2P_OT(r),, RoundingDecimals))
                            Next

                            If SpellingProbability_g2p_ExportValues.Count > 0 Then
                                outputString &= String.Join(" ", SpellingProbability_g2p_ExportValues) & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "ForeignWord" Then
                        If ContainsWordListData = True Then
                            outputString &= ForeignWord & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    'If ColumnOrderProperyInfo(n).Name = "CorrectedSpelling" Then outputString &= CorrectedSpelling & vbTab
                    'If ColumnOrderProperyInfo(n).Name = "CorrectedTranscription" Then outputString &= CorrectedTranscription & vbTab

                    'If ColumnOrderProperyInfo(n).Name = "SAMPA" Then

                    '    'Sampa support has been removed

                    '    'Preparing Sampa string
                    '    'Sampa String
                    '    Dim SampaOutputString As String = ""
                    '    For sampaIndex = 0 To _SAMPA.Count - 1
                    '        If Not _SAMPA(sampaIndex) = "" Then
                    '            SampaOutputString &= _SAMPA(sampaIndex) & " "
                    '        End If
                    '    Next

                    '    'Removes the last blank space
                    '    SampaOutputString = SampaOutputString.TrimEnd(" ")

                    '    'Writing OrthographicForm

                    '    outputString &= SampaOutputString & vbTab

                    'End If

                    If ColumnOrderProperyInfo(n).Name = "ManuallyReveiwedCount" Then outputString &= ManuallyReveiwedCount & vbTab

                    'The following data is output only, and are not read by the input parser
                    If ColumnOrderProperyInfo(n).Name = "IPA" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= String.Join(" ", BuildExtendedIpaArray(,,,,, False, False)) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "ZipfValue" Then outputString &= Rounding(ZipfValue_Word,, 4) & vbTab
                    If ColumnOrderProperyInfo(n).Name = "LetterCount" Then outputString &= OrthographicForm.Length & vbTab

                    If ColumnOrderProperyInfo(n).Name = "GraphemeCount" Then
                        If Sonographs_Letters IsNot Nothing Then
                            outputString &= Sonographs_Letters.Count & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "DiGraphCount" Then
                        If ContainsWordListData = True Then
                            outputString &= DiGraphCount & vbTab
                        Else
                            'Determining if data should be generated
                            If Sonographs_Letters IsNot Nothing And GenerateLackingData = True Then
                                CountComplexGraphemes()
                                outputString &= DiGraphCount & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "TriGraphCount" Then
                        If ContainsWordListData = True Then
                            outputString &= TriGraphCount & vbTab
                        Else
                            'Determining if data should be generated
                            If Sonographs_Letters IsNot Nothing And GenerateLackingData = True Then
                                CountComplexGraphemes()
                                outputString &= TriGraphCount & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "LongGraphemesCount" Then
                        If ContainsWordListData = True Then
                            outputString &= LongGraphemesCount & vbTab
                        Else
                            'Determining if data should be generated
                            If Sonographs_Letters IsNot Nothing And GenerateLackingData = True Then
                                CountComplexGraphemes()
                                outputString &= LongGraphemesCount & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "SyllableCount" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Syllables.Count & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "Tone" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= Tone & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "MainStressSyllable" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= MainStressSyllableIndex & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "SecondaryStressSyllable" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            If SecondaryStressSyllableIndex <> 0 Then
                                outputString &= SecondaryStressSyllableIndex & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PhoneCount" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= LocalPhoneCount & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PhoneCountZero" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= CountPhonemes(True, True) & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PLD1WordCount" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= PLD1WordCount & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "OLD1WordCount" Then
                        If LocalPhoneCount > 0 Or AllowIrrelevantValues = True Then
                            outputString &= OLD1WordCount & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PossiblePoSCount" Then
                        If ContainsWordListData = True Or AllowIrrelevantValues = True Then
                            If AllPossiblePoS.Count > 0 Then
                                outputString &= AllPossiblePoS.Count & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "MostCommonPoS" Then
                        If ContainsWordListData = True Then
                            If MostCommonPoS IsNot Nothing Then
                                outputString &= MostCommonPoS.Item1 & ":" & Rounding(MostCommonPoS.Item2, , RoundingDecimals) & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PossibleLemmaCount" Then
                        If ContainsWordListData = True Or AllowIrrelevantValues = True Then
                            If AllOccurringLemmas.Count > 0 Then
                                outputString &= AllOccurringLemmas.Count & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "MostCommonLemma" Then
                        If ContainsWordListData = True Then
                            If MostCommonLemma IsNot Nothing Then
                                outputString &= MostCommonLemma.Item1 & ":" & Rounding(MostCommonLemma.Item2,, RoundingDecimals) & vbTab
                            Else
                                outputString &= vbTab
                            End If
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "ManualEvaluations" Then outputString &= String.Join("; ", ManualEvaluations).TrimEnd(";") & vbTab
                    If ColumnOrderProperyInfo(n).Name = "ManualEvaluationsCount" Then outputString &= ManualEvaluations.Count & vbTab

                    'Added 2019-08-14
                    If ColumnOrderProperyInfo(n).Name = "OrthographicIsolationPoint" Then
                        If OrthographicIsolationPoint >= 0 Or AllowIrrelevantValues = True Then
                            outputString &= OrthographicIsolationPoint & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    If ColumnOrderProperyInfo(n).Name = "PhoneticIsolationPoint" Then
                        If PhoneticIsolationPoint >= 0 Or AllowIrrelevantValues = True Then
                            outputString &= PhoneticIsolationPoint & vbTab
                        Else
                            outputString &= vbTab
                        End If
                    End If

                    ColumnToWrite += 1
                End If
            End If
        Next


        Return outputString

    End Function




    ''' <summary>
    ''' Generates a word from a string created by the function GenerateFullPhoneticOutputTxtString
    ''' </summary>
    ''' <param name="inputWordString"></param>
    ''' <returns></returns>
    Public Shared Function ParseInputWordString(ByRef inputWordString As String,
                                                     ByRef ColumnOrder As PhoneticTxtStringColumnIndices,
                                                     Optional ByRef CorrectDoubleSpacesInPhoneticForm As Boolean = True,
                                                     Optional ByRef CheckPhonemeValidity As Boolean = True,
                                                     Optional ByRef ValidPhoneticCharacters As List(Of String) = Nothing,
                                                Optional ByRef ColumnLengthList As SortedList(Of String, Integer) = Nothing) As Word


        Dim ContainsInvalidPhoneticCharacter As Boolean = False
        Dim CorrectedDoubleSpacesInPhoneticForm As Boolean = False
        Dim LastAttemptedColumnIndex As Integer = -1
        Dim newWord As New Word
        Dim inputWordStringSplit() As String = inputWordString.Trim(vbTab).Split(vbTab)

        'Making sure there are enough columns in the input array, by adding empty columns if their numbers are too small
        Dim IntendedColumnCount As Integer = ColumnOrder.GetNumberOfColumns
        Dim tempInputWordStringSplit As New List(Of String)
        For n = 0 To inputWordStringSplit.Count - 1
            tempInputWordStringSplit.Add(inputWordStringSplit(n))
        Next
        Do Until tempInputWordStringSplit.Count > IntendedColumnCount
            tempInputWordStringSplit.Add("")
        Loop
        inputWordStringSplit = tempInputWordStringSplit.ToArray

        'Starts reading and storing input data
        Try
            If ColumnOrder.OrthographicForm IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.OrthographicForm
                newWord.OrthographicForm = inputWordStringSplit(ColumnOrder.OrthographicForm).Trim
            End If


            If ColumnOrder.GIL2P_OT_Average IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.GIL2P_OT_Average
                Dim input As String = inputWordStringSplit(ColumnOrder.GIL2P_OT_Average).Trim
                If IsNumeric(input) Then newWord.GIL2P_OT_Average = input
            End If

            If ColumnOrder.GIL2P_OT_Min IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.GIL2P_OT_Min
                Dim input As String = inputWordStringSplit(ColumnOrder.GIL2P_OT_Min).Trim
                If IsNumeric(input) Then newWord.GIL2P_OT_Min = input
            End If

            If ColumnOrder.PIP2G_OT_Average IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.PIP2G_OT_Average
                Dim input As String = inputWordStringSplit(ColumnOrder.PIP2G_OT_Average).Trim
                If IsNumeric(input) Then newWord.PIP2G_OT_Average = input
            End If

            If ColumnOrder.PIP2G_OT_Min IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.PIP2G_OT_Min
                Dim input As String = inputWordStringSplit(ColumnOrder.PIP2G_OT_Min).Trim
                If IsNumeric(input) Then newWord.PIP2G_OT_Min = input
            End If

            If ColumnOrder.G2P_OT_Average IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.G2P_OT_Average
                newWord.G2P_OT_Average = inputWordStringSplit(ColumnOrder.G2P_OT_Average).Trim
            End If

            If ColumnOrder.UpperCase IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.UpperCase
                Dim input As String = inputWordStringSplit(ColumnOrder.UpperCase).Trim
                If IsNumeric(input) Then newWord.ProportionStartingWithUpperCase = input
            End If

            If ColumnOrder.Homographs IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.Homographs

                'Only reading forms if the input string is not empty
                If Not inputWordStringSplit(ColumnOrder.Homographs).Trim = "" Then
                    newWord.LanguageHomographs = New List(Of String)
                    Dim InputForms() As String = inputWordStringSplit(ColumnOrder.Homographs).Trim.Split("|")
                    For CurrentIndex = 0 To InputForms.Length - 1
                        Dim newInputForm As String = InputForms(CurrentIndex).Trim
                        If Not newInputForm = "" Then
                            newWord.LanguageHomographs.Add(newInputForm)
                        End If
                    Next
                End If
            End If

            If ColumnOrder.HomographCount IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.HomographCount
                Dim input As String = inputWordStringSplit(ColumnOrder.HomographCount).Trim
                If IsNumeric(input) Then newWord.LanguageHomographCount = input
                'If newWord.LanguageHomographCount <> newWord.LanguageHomographs.Count Then Errors("The number of actual (language) homographs do not agree with the noted homograph count for word: " & newWord.OrthographicForm)
            End If

            If ColumnOrder.SpecialCharacter IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.SpecialCharacter
                newWord.OrthographicFormContainsSpecialCharacter = inputWordStringSplit(ColumnOrder.SpecialCharacter).Trim
            End If

            If ColumnOrder.RawWordTypeFrequency IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.RawWordTypeFrequency
                Dim input As String = inputWordStringSplit(ColumnOrder.RawWordTypeFrequency).Trim
                If IsNumeric(input) Then newWord.RawWordTypeFrequency = input
            End If

            If ColumnOrder.RawDocumentCount IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.RawDocumentCount
                Dim input As String = inputWordStringSplit(ColumnOrder.RawDocumentCount).Trim
                If IsNumeric(input) Then newWord.RawDocumentCount = input
            End If

            If ColumnOrder.PhoneticForm IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.PhoneticForm

                'Only reading phonetic form if the input string is not empty
                'If Not inputWordStringSplit(ColumnOrder.PhoneticForm).Trim = "" Then

                'ParseInputPhoneticString(inputWordStringSplit(ColumnOrder.PhoneticForm), newWord, ValidPhoneticCharacters,
                '                                 ContainsInvalidPhoneticCharacter, CorrectDoubleSpacesInPhoneticForm,
                '                                 CorrectedDoubleSpacesInPhoneticForm, CheckPhonemeValidity)

                newWord.ParseInputPhoneticString(inputWordStringSplit(ColumnOrder.PhoneticForm), ValidPhoneticCharacters,
                                                 ContainsInvalidPhoneticCharacter, CorrectDoubleSpacesInPhoneticForm,
                                                 CorrectedDoubleSpacesInPhoneticForm, CheckPhonemeValidity)


            End If

            'TemporarySyllabification is not read

            'ReducedTranscription is not read

            If ColumnOrder.PhonotacticType IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.PhonotacticType
                newWord.PhonotacticType = inputWordStringSplit(ColumnOrder.PhonotacticType).Trim
            End If


            If ColumnOrder.SSPP_Average IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.SSPP_Average
                Dim input As String = inputWordStringSplit(ColumnOrder.SSPP_Average).Trim
                If IsNumeric(input) Then newWord.SSPP_Average = input
            End If

            If ColumnOrder.SSPP_Min IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.SSPP_Min
                Dim input As String = inputWordStringSplit(ColumnOrder.SSPP_Min).Trim
                If IsNumeric(input) Then newWord.SSPP_Min = input
            End If

            If ColumnOrder.PSP_Sum IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.PSP_Sum
                Dim input As String = inputWordStringSplit(ColumnOrder.PSP_Sum).Trim
                If IsNumeric(input) Then newWord.PSP_Sum = input
            End If

            If ColumnOrder.PSBP_Sum IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.PSBP_Sum
                Dim input As String = inputWordStringSplit(ColumnOrder.PSBP_Sum).Trim
                If IsNumeric(input) Then newWord.PSBP_Sum = input
            End If

            If ColumnOrder.S_PSP_Average IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.S_PSP_Average
                Dim input As String = inputWordStringSplit(ColumnOrder.S_PSP_Average).Trim
                If IsNumeric(input) Then newWord.S_PSP_Average = input
            End If

            If ColumnOrder.S_PSBP_Average IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.S_PSBP_Average
                Dim input As String = inputWordStringSplit(ColumnOrder.S_PSBP_Average).Trim
                If IsNumeric(input) Then newWord.S_PSBP_Average = input
            End If

            If ColumnOrder.Homophones IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.Homophones

                'Only reading forms if the input string is not empty
                If Not inputWordStringSplit(ColumnOrder.Homophones).Trim = "" Then
                    newWord.LanguageHomophones = New List(Of String)
                    Dim InputForms() As String = inputWordStringSplit(ColumnOrder.Homophones).Trim.Split("|")
                    For CurrentIndex = 0 To InputForms.Length - 1
                        Dim newInputForm As String = InputForms(CurrentIndex).Trim
                        If Not newInputForm = "" Then
                            newWord.LanguageHomophones.Add(newInputForm)
                        End If
                    Next
                End If
            End If

            If ColumnOrder.HomophoneCount IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.HomophoneCount
                Dim input As String = inputWordStringSplit(ColumnOrder.HomophoneCount).Trim
                If IsNumeric(input) Then newWord.LanguageHomophoneCount = input
                'If newWord.LanguageHomophoneCount <> newWord.LanguageHomophones.Count Then Errors("The number of actual (language) homophones do not agree with the noted homophone count for word: " & newWord.OrthographicForm)
            End If

            If ColumnOrder.PNDP IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.PNDP
                If inputWordStringSplit(ColumnOrder.PNDP).Trim <> "" Then
                    Dim input As String = inputWordStringSplit(ColumnOrder.PNDP).Trim
                    If IsNumeric(input) Then newWord.FWPN_DensityProbability = input
                End If
            End If

            If ColumnOrder.PLD1Transcriptions IsNot Nothing Then

                'Reading PLD1Transcriptions
                LastAttemptedColumnIndex = ColumnOrder.PLD1Transcriptions
                'Only reads PLD1Transcriptions if the string is not empty
                If Not inputWordStringSplit(ColumnOrder.PLD1Transcriptions).Trim = "" Then

                    Dim AllPLD1Transcriptions() As String = inputWordStringSplit(ColumnOrder.PLD1Transcriptions).Trim.Split("|")
                    For i = 0 To AllPLD1Transcriptions.Length - 1
                        newWord.PLD1Transcriptions.Add(AllPLD1Transcriptions(i))
                    Next

                    'Reporting erroneous PLD1 transcription arrays of length = 1
                    If newWord.PLD1Transcriptions.Count = 1 Then MsgBox("The input word " & newWord.OrthographicForm & " has an erroneous PLD1 transcription array. Array length should never be 1!" & vbCrLf & "The PLD1 input string look as follows:" & inputWordStringSplit(ColumnOrder.PLD1Transcriptions))

                    'Setting PLD1 Word count
                    If newWord.PLD1Transcriptions.Count = 0 Then
                        newWord.PLD1WordCount = 0
                    Else
                        newWord.PLD1WordCount = newWord.PLD1Transcriptions.Count - 1 '-1 is used since the first transcription in the PLD1Transcriptions array is the current member word
                    End If

                End If
            End If

            If ColumnOrder.ONDP IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.ONDP
                If inputWordStringSplit(ColumnOrder.ONDP).Trim <> "" Then
                    Dim input As String = inputWordStringSplit(ColumnOrder.ONDP).Trim
                    If IsNumeric(input) Then newWord.FWON_DensityProbability = input
                End If
            End If

            If ColumnOrder.OLD1Spellings IsNot Nothing Then

                'Reading OLD1Transcriptions
                LastAttemptedColumnIndex = ColumnOrder.OLD1Spellings
                'Only reads OLD1Transcriptions if the string is not empty
                If Not inputWordStringSplit(ColumnOrder.OLD1Spellings).Trim = "" Then

                    Dim AllOLD1Spellings() As String = inputWordStringSplit(ColumnOrder.OLD1Spellings).Trim.Split("|")
                    For i = 0 To AllOLD1Spellings.Length - 1
                        newWord.OLD1Spellings.Add(AllOLD1Spellings(i))
                    Next

                    'Reporting erroneous OLD1 transcription arrays of length = 1
                    If newWord.OLD1Spellings.Count = 1 Then MsgBox("The input word " & newWord.OrthographicForm & " has an erroneous OLD1 transcription array. Array length should never be 1!" & vbCrLf & "The OLD1 input string look as follows:" & inputWordStringSplit(ColumnOrder.OLD1Spellings))

                    'Setting OLD1 Word count
                    If newWord.OLD1Spellings.Count = 0 Then
                        newWord.OLD1WordCount = 0
                    Else
                        newWord.OLD1WordCount = newWord.OLD1Spellings.Count - 1 '-1 is used since the first spelling in the OLD1Spellings array is the current member word
                    End If

                End If
            End If


            If ColumnOrder.PLDx_Average IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.PLDx_Average
                If inputWordStringSplit(ColumnOrder.PLDx_Average).Trim <> "" Then
                    Dim input As String = inputWordStringSplit(ColumnOrder.PLDx_Average).Trim
                    If IsNumeric(input) Then newWord.PLDx_Average = input
                End If
            End If

            If ColumnOrder.OLDx_Average IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.OLDx_Average
                If inputWordStringSplit(ColumnOrder.OLDx_Average).Trim <> "" Then
                    Dim input As String = inputWordStringSplit(ColumnOrder.OLDx_Average).Trim
                    If IsNumeric(input) Then newWord.OLDx_Average = input
                End If
            End If

            If ColumnOrder.PLDx_Neighbors IsNot Nothing Then

                'Reading PLDx_Neighbors
                LastAttemptedColumnIndex = ColumnOrder.PLDx_Neighbors
                'Only reads PLDx_Neighbors if the string is not empty
                If Not inputWordStringSplit(ColumnOrder.PLDx_Neighbors).Trim = "" Then

                    Dim AllPldxData() As String = inputWordStringSplit(ColumnOrder.PLDx_Neighbors).Trim.Split("|")
                    For i = 0 To AllPldxData.Length - 1
                        Dim CurrentSplit() As String = AllPldxData(i).Split(":")
                        If CurrentSplit.Length > 2 Then
                            newWord.PLDxData.Add(New Tuple(Of Integer, String, Single)(CurrentSplit(0), CurrentSplit(1), CurrentSplit(2)))
                        End If
                    Next
                End If
            End If


            If ColumnOrder.OLDx_Neighbors IsNot Nothing Then

                'Reading OLDx_Neighbors
                LastAttemptedColumnIndex = ColumnOrder.OLDx_Neighbors
                'Only reads OLDx_Neighbors if the string is not empty
                If Not inputWordStringSplit(ColumnOrder.OLDx_Neighbors).Trim = "" Then

                    Dim AllOldxData() As String = inputWordStringSplit(ColumnOrder.OLDx_Neighbors).Trim.Split("|")
                    For i = 0 To AllOldxData.Length - 1
                        Dim CurrentSplit() As String = AllOldxData(i).Split(":")
                        If CurrentSplit.Length > 2 Then
                            newWord.OLDxData.Add(New Tuple(Of Integer, String, Single)(CurrentSplit(0), CurrentSplit(1), CurrentSplit(2)))
                        End If
                    Next
                End If
            End If

            If ColumnOrder.Sonographs IsNot Nothing Then

                LastAttemptedColumnIndex = ColumnOrder.Sonographs

                'Reading graphemes, only if the input string is not empty
                If Not inputWordStringSplit(ColumnOrder.Sonographs).Trim = "" Then
                    Dim AllSonographs() As String = inputWordStringSplit(ColumnOrder.Sonographs).Trim.Split("|")
                    newWord.Sonographs_Letters = New List(Of String)
                    newWord.Sonographs_Pronunciation = New List(Of String)
                    For Grapheme = 0 To AllSonographs.Length - 1

                        Dim SonographsSplit() As String = AllSonographs(Grapheme).Trim.Split("-")
                        newWord.Sonographs_Letters.Add(SonographsSplit(0).Trim)

                        'Adds phoneme blocks only if they exist
                        If SonographsSplit.Length > 1 Then
                            newWord.Sonographs_Pronunciation.Add(SonographsSplit(1).Trim)
                        End If

                    Next
                End If
            End If

            If ColumnOrder.AllPoS IsNot Nothing Then

                'Reading Parts of speech
                LastAttemptedColumnIndex = ColumnOrder.AllPoS
                'Only reads PoS if the string is not empty
                If Not inputWordStringSplit(ColumnOrder.AllPoS).Trim = "" Then

                    Dim AllPoS() As String

                    'Allowing for semicolon delimiter used in previous formats
                    If inputWordStringSplit(ColumnOrder.AllPoS).Trim.Contains(";") Then
                        'Parsing using semicolon delimiter
                        AllPoS = inputWordStringSplit(ColumnOrder.AllPoS).Trim.Split(";")
                    Else
                        'Parsing using | delimiter
                        AllPoS = inputWordStringSplit(ColumnOrder.AllPoS).Trim.Split("|")
                    End If

                    Dim AlreadyAddedPoSs As New SortedSet(Of String)
                    For PoS = 0 To AllPoS.Length - 1
                        Dim CurrentPoSSplit() As String = AllPoS(PoS).Trim.Split(":")

                        Dim Temp_PoS As String = CurrentPoSSplit(0).Trim

                        If Not AlreadyAddedPoSs.Contains(Temp_PoS) Then
                            AlreadyAddedPoSs.Add(Temp_PoS)

                            If CurrentPoSSplit.Length = 1 Then
                                newWord.AllPossiblePoS.Add(New Tuple(Of String, Double)(Temp_PoS, 0))
                            ElseIf CurrentPoSSplit.Length > 1 Then
                                newWord.AllPossiblePoS.Add(New Tuple(Of String, Double)(Temp_PoS, CurrentPoSSplit(1).Trim))
                            End If

                        Else
                            MsgBox("Detected Error In a Part-Of-Speech String. Duplicate Parts-Of-speech For word " & newWord.OrthographicForm & vbCr &
                                           "Please go back To the excel file And check For errors And Then reload the material!")
                        End If
                    Next

                    'Setting most common PoS
                    If newWord.AllPossiblePoS.Count > 0 Then
                        'newWord.MostCommonPoS = New StringDoubleCombination
                        Dim HigestPoSValue As Double = 0
                        Dim HigestPoSIndex As Integer = 0
                        For PoSIndex = 0 To newWord.AllPossiblePoS.Count - 1
                            If newWord.AllPossiblePoS(PoSIndex).Item2 > HigestPoSValue Then
                                HigestPoSValue = newWord.AllPossiblePoS(PoSIndex).Item2
                                HigestPoSIndex = PoSIndex
                            End If
                        Next
                        newWord.MostCommonPoS = newWord.AllPossiblePoS(HigestPoSIndex)
                    End If
                End If
            End If

            If ColumnOrder.AllLemmas IsNot Nothing Then

                'Reading lemmas
                LastAttemptedColumnIndex = ColumnOrder.AllLemmas
                'Only reads AllLemmas if the string is not empty
                If Not inputWordStringSplit(ColumnOrder.AllLemmas).Trim = "" Then

                    Dim AllLemmas() As String = inputWordStringSplit(ColumnOrder.AllLemmas).Trim.Split("|")
                    Dim AlreadyAddedLemmas As New SortedSet(Of String)
                    For Lemma = 0 To AllLemmas.Length - 1
                        Dim CurrentLemmaSplit() As String = AllLemmas(Lemma).Trim.Split(":")
                        Dim tempLemma As String = CurrentLemmaSplit(0).Trim

                        If Not AlreadyAddedLemmas.Contains(tempLemma) Then
                            AlreadyAddedLemmas.Add(tempLemma)
                            If CurrentLemmaSplit.Length = 1 Then
                                newWord.AllOccurringLemmas.Add(New Tuple(Of String, Double)(tempLemma, 0))
                            ElseIf CurrentLemmaSplit.Length > 1 Then
                                newWord.AllOccurringLemmas.Add(New Tuple(Of String, Double)(tempLemma, CurrentLemmaSplit(1).Trim))
                            End If
                        Else
                            MsgBox("Detected Error In a lemma input String. Duplicate lemmas for word " & newWord.OrthographicForm & vbCr &
                                       "Please go back to the input .txt file And check for errors and then reload the material!")
                        End If
                    Next

                    'Setting most common lemma
                    If newWord.AllOccurringLemmas.Count > 0 Then
                        'newWord.MostCommonLemma = New StringDoubleCombination
                        Dim HigestLemmaValue As Double = 0
                        Dim HigestLemmaIndex As Integer = 0
                        For LemmaIndex = 0 To newWord.AllOccurringLemmas.Count - 1
                            If newWord.AllOccurringLemmas(LemmaIndex).Item2 > HigestLemmaValue Then
                                HigestLemmaValue = newWord.AllOccurringLemmas(LemmaIndex).Item2
                                HigestLemmaIndex = LemmaIndex
                            End If
                        Next
                        newWord.MostCommonLemma = newWord.AllOccurringLemmas(HigestLemmaIndex)
                    End If
                End If
            End If

            If ColumnOrder.NumberOfSenses IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.NumberOfSenses
                Dim input As String = inputWordStringSplit(ColumnOrder.NumberOfSenses).Trim
                If IsNumeric(input) Then newWord.NumberOfSenses = input
            End If

            If ColumnOrder.Abbreviation IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.Abbreviation
                newWord.Abbreviation = inputWordStringSplit(ColumnOrder.Abbreviation).Trim
            End If

            If ColumnOrder.Acronym IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.Acronym
                newWord.Acronym = inputWordStringSplit(ColumnOrder.Acronym).Trim
            End If



            If ColumnOrder.AllPossibleSenses IsNot Nothing Then

                'Reading senses
                LastAttemptedColumnIndex = ColumnOrder.AllPossibleSenses
                'Only reads AllPossibleSenses if the string is not empty
                If Not inputWordStringSplit(ColumnOrder.AllPossibleSenses).Trim = "" Then

                    Dim AllSenses() As String = inputWordStringSplit(ColumnOrder.AllPossibleSenses).Trim.Split("|")
                    Dim AlreadyAddedSenses As New SortedSet(Of String)
                    For Sense = 0 To AllSenses.Length - 1
                        Dim CurrentSense As String = AllSenses(Sense).Trim
                        If Not AlreadyAddedSenses.Contains(CurrentSense) Then
                            AlreadyAddedSenses.Add(CurrentSense)
                            newWord.AllPossibleSenses.Add(CurrentSense)
                        Else
                            MsgBox("Detected Error in a possible senses input string. Duplicate senses for word " & newWord.OrthographicForm & vbCr &
                                       "Please go back to the input .txt file And check for errors and then reload the material!")
                        End If
                    Next

                    'Setting Senses count
                    newWord.NumberOfSenses = newWord.AllPossibleSenses.Count
                End If
            End If

            If ColumnOrder.SSPP IsNot Nothing Then

                LastAttemptedColumnIndex = ColumnOrder.SSPP

                'Reading PhonoTacticProbability data, only if the input string is not empty
                If Not inputWordStringSplit(ColumnOrder.SSPP).Trim = "" Then

                    newWord.PP_Phonemes = New List(Of String)
                    newWord.PP = New List(Of Double)

                    Dim PhonoTacticCombination() As String = inputWordStringSplit(ColumnOrder.SSPP).Trim.Split(" ")
                    For i = 0 To PhonoTacticCombination.Count - 2 Step 2
                        newWord.PP_Phonemes.Add(PhonoTacticCombination(i).TrimStart("[").TrimEnd("]"))
                        newWord.PP.Add(PhonoTacticCombination(i + 1).Trim)
                    Next

                    'Adding also the word end string
                    newWord.PP_Phonemes.Add(PhonoTacticCombination(PhonoTacticCombination.Count - 1))

                End If
            End If

            If ColumnOrder.PSP IsNot Nothing Then

                LastAttemptedColumnIndex = ColumnOrder.PSP

                'Reading PSP data, only if the input string is not empty
                If Not inputWordStringSplit(ColumnOrder.PSP).Trim = "" Then

                    newWord.PSP_Phonemes = New List(Of String)
                    newWord.PSP = New List(Of Double)

                    Dim PhonoTacticCombination() As String = inputWordStringSplit(ColumnOrder.PSP).Trim.Split(" ")
                    For i = 0 To PhonoTacticCombination.Count - 2 Step 2
                        newWord.PSP_Phonemes.Add(PhonoTacticCombination(i).TrimStart("[").TrimEnd("]"))
                        newWord.PSP.Add(PhonoTacticCombination(i + 1).Trim)
                    Next

                End If
            End If

            If ColumnOrder.PSBP IsNot Nothing Then

                LastAttemptedColumnIndex = ColumnOrder.PSBP

                'Reading PSBP data, only if the input string is not empty
                If Not inputWordStringSplit(ColumnOrder.PSBP).Trim = "" Then

                    newWord.PSBP_Phonemes = New List(Of String)
                    newWord.PSBP = New List(Of Double)

                    Dim PhonoTacticCombinationString As String = inputWordStringSplit(ColumnOrder.PSBP).Trim
                    'Replacing "[" and "]" with vbTabs, which can be used to segment into an array
                    PhonoTacticCombinationString = PhonoTacticCombinationString.Replace("[", vbTab)
                    PhonoTacticCombinationString = PhonoTacticCombinationString.Replace("]", vbTab)
                    Dim PhonoTacticCombination() As String = PhonoTacticCombinationString.Trim.Split(vbTab)
                    For i = 0 To PhonoTacticCombination.Count - 2 Step 2
                        Try
                            newWord.PSBP_Phonemes.Add(PhonoTacticCombination(i).Trim)
                            newWord.PSBP.Add(PhonoTacticCombination(i + 1).Trim)
                        Catch ex As Exception
                            MsgBox(ex.ToString)
                        End Try

                    Next

                End If
            End If

            If ColumnOrder.S_PSP IsNot Nothing Then

                LastAttemptedColumnIndex = ColumnOrder.S_PSP

                'Reading S_PSP data, only if the input string is not empty
                If Not inputWordStringSplit(ColumnOrder.S_PSP).Trim = "" Then

                    newWord.S_PSP_Phonemes = New List(Of String)
                    newWord.S_PSP = New List(Of Double)

                    Dim PhonoTacticCombination() As String = inputWordStringSplit(ColumnOrder.S_PSP).Trim.Split(" ")
                    For i = 0 To PhonoTacticCombination.Count - 2 Step 2
                        newWord.S_PSP_Phonemes.Add(PhonoTacticCombination(i).TrimStart("[").TrimEnd("]"))
                        newWord.S_PSP.Add(PhonoTacticCombination(i + 1).Trim)
                    Next

                End If
            End If

            If ColumnOrder.S_PSBP IsNot Nothing Then

                LastAttemptedColumnIndex = ColumnOrder.S_PSBP

                'Reading S_PSBP data, only if the input string is not empty
                If Not inputWordStringSplit(ColumnOrder.S_PSBP).Trim = "" Then

                    newWord.S_PSBP_Phonemes = New List(Of String)
                    newWord.S_PSBP = New List(Of Double)

                    Dim PhonoTacticCombinationString As String = inputWordStringSplit(ColumnOrder.S_PSBP).Trim
                    'Replacing "[" and "]" with vbTabs, which can be used to segment into an array
                    PhonoTacticCombinationString = PhonoTacticCombinationString.Replace("[", vbTab)
                    PhonoTacticCombinationString = PhonoTacticCombinationString.Replace("]", vbTab)
                    Dim PhonoTacticCombination() As String = PhonoTacticCombinationString.Trim.Split(vbTab)
                    For i = 0 To PhonoTacticCombination.Count - 2 Step 2
                        Try
                            newWord.S_PSBP_Phonemes.Add(PhonoTacticCombination(i).Trim)
                            newWord.S_PSBP.Add(PhonoTacticCombination(i + 1).Trim)
                        Catch ex As Exception
                            MsgBox(ex.ToString)
                        End Try

                    Next

                End If
            End If



            If ColumnOrder.GIL2P_OT IsNot Nothing Then

                LastAttemptedColumnIndex = ColumnOrder.GIL2P_OT

                'Reading GIL2P_OT data, only if the input string is not empty
                If Not inputWordStringSplit(ColumnOrder.GIL2P_OT).Trim = "" Then
                    Dim AllGIL2P_OT() As String = inputWordStringSplit(ColumnOrder.GIL2P_OT).Trim.Split(" ")
                    For n = 0 To AllGIL2P_OT.Length - 1
                        newWord.GIL2P_OT.Add(AllGIL2P_OT(n).Trim)
                    Next
                End If
            End If

            If ColumnOrder.PIP2G_OT IsNot Nothing Then

                LastAttemptedColumnIndex = ColumnOrder.PIP2G_OT

                'Reading PIP2G_OT data, only if the input string is not empty
                If Not inputWordStringSplit(ColumnOrder.PIP2G_OT).Trim = "" Then
                    Dim AllPIP2G_OT() As String = inputWordStringSplit(ColumnOrder.PIP2G_OT).Trim.Split(" ")
                    For n = 0 To AllPIP2G_OT.Length - 1
                        newWord.PIP2G_OT.Add(AllPIP2G_OT(n).Trim)
                    Next
                End If
            End If

            If ColumnOrder.G2P_OT IsNot Nothing Then

                LastAttemptedColumnIndex = ColumnOrder.G2P_OT

                'Reading G2P_OT data, only if the input string is not empty
                If Not inputWordStringSplit(ColumnOrder.G2P_OT).Trim = "" Then
                    Dim AllG2P_OT() As String = inputWordStringSplit(ColumnOrder.G2P_OT).Trim.Split(" ")
                    For n = 0 To AllG2P_OT.Length - 1
                        newWord.G2P_OT.Add(AllG2P_OT(n).Trim)
                    Next
                End If
            End If

            If ColumnOrder.ForeignWord IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.ForeignWord
                newWord.ForeignWord = inputWordStringSplit(ColumnOrder.ForeignWord).Trim
            End If

            'If ColumnOrder.CorrectedSpelling IsNot Nothing Then
            '    LastAttemptedColumnIndex = ColumnOrder.CorrectedSpelling
            '    newWord.CorrectedSpelling = inputWordStringSplit(ColumnOrder.CorrectedSpelling).Trim
            'End If

            'If ColumnOrder.CorrectedTranscription IsNot Nothing Then
            '    LastAttemptedColumnIndex = ColumnOrder.CorrectedTranscription
            '    newWord.CorrectedTranscription = inputWordStringSplit(ColumnOrder.CorrectedTranscription).Trim
            'End If

            'SAMPA support has been removed
            'If ColumnOrder.SAMPA IsNot Nothing Then
            '    LastAttemptedColumnIndex = ColumnOrder.SAMPA

            '    'Only reading SAMPA forms if the input string is not empty
            '    If Not inputWordStringSplit(ColumnOrder.SAMPA).Trim = "" Then
            '        Dim SampaForms() As String = inputWordStringSplit(ColumnOrder.SAMPA).Trim.Split(" ")
            '        newWord._SAMPA = New List(Of String)
            '        For SampaFormIndex = 0 To SampaForms.Length - 1
            '            Dim newSampaForm As String = SampaForms(SampaFormIndex).Trim
            '            If Not newSampaForm = "" Then
            '                newWord._SAMPA.Add(newSampaForm)
            '            End If
            '        Next
            '    End If
            'End If


            If ColumnOrder.ManuallyReveiwedCount IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.ManuallyReveiwedCount
                Dim input As String = inputWordStringSplit(ColumnOrder.ManuallyReveiwedCount).Trim
                If IsNumeric(input) Then newWord.ManuallyReveiwedCount = input
            End If

            'Addition 2020-05-23
            If ColumnOrder.ManualEvaluations IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.ManualEvaluations
                Dim input() As String = inputWordStringSplit(ColumnOrder.ManualEvaluations).Trim.Split(",")
                For n = 0 To input.Length - 1
                    If Not input(n).Trim = "" Then newWord.ManualEvaluations.Add(input(n).Trim)
                Next
            End If


            'Added 2019-08-14
            If ColumnOrder.OrthographicIsolationPoint IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.OrthographicIsolationPoint
                If inputWordStringSplit(ColumnOrder.OrthographicIsolationPoint).Trim <> "" Then
                    Dim input As String = inputWordStringSplit(ColumnOrder.OrthographicIsolationPoint).Trim
                    If IsNumeric(input) Then newWord.OrthographicIsolationPoint = input
                End If
            End If

            If ColumnOrder.PhoneticIsolationPoint IsNot Nothing Then
                LastAttemptedColumnIndex = ColumnOrder.PhoneticIsolationPoint
                If inputWordStringSplit(ColumnOrder.PhoneticIsolationPoint).Trim <> "" Then
                    Dim input As String = inputWordStringSplit(ColumnOrder.PhoneticIsolationPoint).Trim
                    If IsNumeric(input) Then newWord.PhoneticIsolationPoint = input
                End If
            End If

            'Logging parsing errors only if the word contains a phonetic transcription
            If newWord.Syllables.Count > 0 Then
                If ContainsInvalidPhoneticCharacter = True Then SendInfoToLog(newWord.OrthographicForm & vbTab & String.Join(" ", newWord.BuildExtendedIpaArray),
                                                                                       "InputFileContainsInvalidPhoneticCharacters")
                If CorrectedDoubleSpacesInPhoneticForm = True Then SendInfoToLog(newWord.OrthographicForm & vbTab & String.Join(" ", newWord.BuildExtendedIpaArray),
                                                                               "CorrectedDoubleSpacesInPhoneticForm")
            End If

            'Counting column lengths
            If ColumnLengthList IsNot Nothing = True Then

                DetectMaxColumnLengths(inputWordStringSplit, ColumnOrder, ColumnLengthList)

            End If


        Catch ex As Exception
            Dim tempOrthForm As String = "Not Found"

            Dim ColumnOrderProperyInfo() As System.Reflection.PropertyInfo = GetType(PhoneticTxtStringColumnIndices).GetProperties
            Dim ColumnName As String = ColumnOrderProperyInfo(LastAttemptedColumnIndex).Name

            If newWord.OrthographicForm <> "" Then tempOrthForm = newWord.OrthographicForm
            MsgBox("An error occurred while parsing the input word file." & vbCr & vbCr &
                       "The error occurred when attempting to parse the following data: " & vbCr &
                       "Orthographic form: " & tempOrthForm & vbCr &
                       "Column index: " & LastAttemptedColumnIndex & vbCr &
                       "Column heading: " & ColumnName & vbCr & vbCr &
                       "Make sure the specified column indices (zero-based) match the input file." & vbCr & vbCr & ex.ToString)
        End Try

        Return newWord

    End Function


    ''' <summary>
    ''' Parses the stardard IPA phonetic form into a syllable structure in the current word
    ''' </summary>
    ''' <param name="PhoneticInputString">The input IPA form. (Phonetic items should be space delimited. Syllable boundary marker is [.])</param>
    ''' <param name="ValidPhoneticCharacters"></param>
    ''' <param name="ContainsInvalidPhoneticCharacter"></param>
    ''' <param name="CorrectDoubleSpacesInPhoneticForm"></param>
    ''' <param name="CorrectedDoubleSpacesInPhoneticForm"></param>
    ''' <param name="CheckPhonemeValidity"></param>
    Public Sub ParseInputPhoneticString(ByRef PhoneticInputString As String,
                                             ByRef ValidPhoneticCharacters As List(Of String),
                                            ByRef ContainsInvalidPhoneticCharacter As Boolean,
                                            ByRef CorrectDoubleSpacesInPhoneticForm As Boolean,
                                            ByRef CorrectedDoubleSpacesInPhoneticForm As Boolean,
                                            ByRef CheckPhonemeValidity As Boolean)




        'Only reading phonetic form if the input string is not empty
        If Not PhoneticInputString.Trim = "" Then

            'Reading the extended syllable array and parsing it to a syllable
            Dim ExtendedIpaArraySyllableSplit() As String = {}

            ExtendedIpaArraySyllableSplit = PhoneticInputString.Trim(".").Trim.Split(".")

            'Replacing any double spaces in the input PhoneticForm with single spaces
            If CorrectDoubleSpacesInPhoneticForm = True Then
                For s = 0 To ExtendedIpaArraySyllableSplit.Count - 1
                    If ExtendedIpaArraySyllableSplit(s).Contains("  ") Then
                        'Correct the double space (such should not occur)
                        ExtendedIpaArraySyllableSplit(s) = ExtendedIpaArraySyllableSplit(s).Replace("  ", " ")
                        CorrectedDoubleSpacesInPhoneticForm = True
                    End If
                Next
            End If

            'Checking that all phonetic characters are valid
            If CheckPhonemeValidity = True And ValidPhoneticCharacters IsNot Nothing Then
                For s = 0 To ExtendedIpaArraySyllableSplit.Length - 1
                    Dim SyllSplit() As String = ExtendedIpaArraySyllableSplit(s).Trim.Split(" ")
                    For p = 0 To SyllSplit.Length - 1
                        If Not ValidPhoneticCharacters.Contains(SyllSplit(p)) Then
                            ContainsInvalidPhoneticCharacter = True
                        End If
                    Next
                Next
            End If


            'Reading suprasegmentals (Tone and index of primary and secondary stress) from the ExtendedIpaArray
            'Setting default values
            Syllables = New ListOfSyllables
            Syllables.Tone = 0
            Syllables.MainStressSyllableIndex = 0
            Syllables.SecondaryStressSyllableIndex = 0

            'Reading values
            For syllable = 0 To ExtendedIpaArraySyllableSplit.Count - 1

                'Detecting primary stress and tone 1
                If ExtendedIpaArraySyllableSplit(syllable).Contains(IpaMainStress) Then
                    Syllables.Tone = 1
                    Syllables.MainStressSyllableIndex = syllable + 1
                End If

                'Detecting primary stress and tone 2
                If ExtendedIpaArraySyllableSplit(syllable).Contains(IpaMainSwedishAccent2) Then
                    Syllables.Tone = 2
                    Syllables.MainStressSyllableIndex = syllable + 1
                End If

                'Detecting secondary stress
                If ExtendedIpaArraySyllableSplit(syllable).Contains(IpaSecondaryStress) Then
                    Syllables.SecondaryStressSyllableIndex = syllable + 1
                End If
            Next

            'Also copying the suprasegmentals to the word variebles
            Tone = Syllables.Tone
            MainStressSyllableIndex = Syllables.MainStressSyllableIndex
            SecondaryStressSyllableIndex = Syllables.SecondaryStressSyllableIndex


            For syllable = 0 To ExtendedIpaArraySyllableSplit.Count - 1

                Dim newSyllable As New Word.Syllable

                'Sets syllable suprasegmentals
                If (syllable = Syllables.MainStressSyllableIndex - 1 And Syllables.MainStressSyllableIndex <> 0) Then
                    newSyllable.IsStressed = True
                End If
                If (syllable = Syllables.SecondaryStressSyllableIndex - 1 And Syllables.SecondaryStressSyllableIndex <> 0) Then
                    newSyllable.IsStressed = True
                    newSyllable.CarriesSecondaryStress = True
                End If

                Dim SyllableArraySplit() As String = ExtendedIpaArraySyllableSplit(syllable).Trim(" ").Split(" ")
                For phoneme = 0 To SyllableArraySplit.Count - 1
                    'Adding only phoneme characters
                    If Not AllSuprasegmentalIPACharacters.Contains(SyllableArraySplit(phoneme)) Then
                        newSyllable.Phonemes.Add(SyllableArraySplit(phoneme).Trim)
                    End If
                Next

                'Detecting and removing ambigous syllable markers
                If newSyllable.Phonemes(0).Contains(AmbiguosOnsetMarker) Then
                    newSyllable.AmbigousOnset = True
                    newSyllable.Phonemes(0) = newSyllable.Phonemes(0).Replace(AmbiguosOnsetMarker, "")
                End If
                If newSyllable.Phonemes(0).Contains(AmbiguosCodaMarker) Then
                    newSyllable.AmbigousCoda = True
                    newSyllable.Phonemes(0) = newSyllable.Phonemes(0).Replace(AmbiguosCodaMarker, "")
                End If

                Syllables.Add(newSyllable)
            Next
        End If


    End Sub


    ''' <summary>
    ''' Detects the highest number of (unparsed) charachters in each ARC-list column.
    ''' </summary>
    ''' <param name="inputWordStringSplit"></param>
    ''' <param name="ColumnOrder"></param>
    ''' <param name="ColumnLengthList"></param>
    Private Shared Sub DetectMaxColumnLengths(ByRef inputWordStringSplit As String(),
                                            ByRef ColumnOrder As PhoneticTxtStringColumnIndices,
                                            ByRef ColumnLengthList As SortedList(Of String, Integer))

        Dim ColumnOrderProperyInfo() As System.Reflection.PropertyInfo = GetType(PhoneticTxtStringColumnIndices).GetProperties
        Dim WriteColumn As Integer = 0
        For n = 0 To ColumnOrderProperyInfo.Length - 1
            If GetType(PhoneticTxtStringColumnIndices).GetProperty(ColumnOrderProperyInfo(n).Name).GetValue(ColumnOrder) IsNot Nothing Then

                ColumnLengthList(ColumnOrderProperyInfo(n).Name) = Math.Max(ColumnLengthList(ColumnOrderProperyInfo(n).Name), inputWordStringSplit(WriteColumn).Length)
                WriteColumn += 1

            End If
        Next

    End Sub


#End Region



End Class







