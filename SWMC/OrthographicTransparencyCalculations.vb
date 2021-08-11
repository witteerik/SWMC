
Imports System.IO


''' <summary>
''' This class contains parameters used for the phonetic transcription-to-spelling (p2g) parsing provided by the functionality 
''' in the methods GetSonographs in the WordGroup class and Generate_p2g_Data in the Word class.
''' </summary>
Public Class p2gParameters

    ''' <summary>
    ''' 
    ''' </summary>
    Public PhoneToGraphemesDictionary As PhoneToGraphemes

    ''' <summary>
    ''' 
    ''' </summary>
    Public ErrorDictionary As Dictionary(Of String, WordGroup.CountExamples)

    ''' <summary>
    ''' 
    ''' </summary>
    Public ListOfGraphemes As List(Of String)

        ''' <summary>
        ''' Select true to use a file containing temporary changes to the orthographic form: One type of change on each line: [present orthographoc form] tab [temporary orthagraphic form]. Also small part of words may be exchanged (e.g. x -> ks). The temporary spelling will be used instead of the original spelling in the matching of phonetic and orthographic forms.
        ''' The order of similar forms should be orderred from the longest to the shortest [present orthographoc form]. For example a 2 can be changed to "two", but a 22 need to be put before 2 in the input list in order to avoid 22 two be changed to "twotwo".
        ''' </summary>
        Public UseTemporarySpellingChange As Boolean

        ''' <summary>
        ''' 
        ''' </summary>
        Public TemporarySpellingChangeListArray() As String

        ''' <summary>
        ''' 
        ''' </summary>
        Public WarnForIllegalCharactersInSpelling As Boolean

        ''' <summary>
        ''' 
        ''' </summary>
        Public LengthSensitiveLookUpPhonemes As Boolean

        ''' <summary>
        ''' If set to true, and if specified with the marker PDCS for the phoneme in the rules file, and if matching fails, an attempt will be made to match the forms with the specified phoneme doubled (since it may have been deleted in the phonetic form, such as may happen in compounds where two juxtaposed stems end and start with the same sound). However this is never attempted for the first phoneme in a word.
        ''' </summary>
        Public UsePDCS_Addition As Boolean

        ''' <summary>
        ''' If set to True (and if UsePDCS_Addition is set to True), PDCS addition will be used also for the last phoneme in a word, otherwise PDCS addition will only be attempted on the second through the second phoneme from the end of the word.
        ''' </summary>
        Public UsePDCS_AdditionOnLastPhoneme As Boolean

        ''' <summary>
        ''' 
        ''' </summary>
        Public UsePDCS_AdditionsWithoutLength As Boolean

        ''' <summary>
        ''' N.B. A spell check should always be performed before using PDS addition. Never used on the first phoneme. Optional use on the last phoneme.
        ''' </summary>
        Public UsePDS_Addition As Boolean

        ''' <summary>
        ''' 
        ''' </summary>
        Public UsePDS_AdditionOnLastPhoneme As Boolean

        ''' <summary>
        ''' 
        ''' </summary>
        Public UsePhonemeReplacement As Boolean

        ''' <summary>
        ''' 
        ''' </summary>
        Public DoRealPhoneticFormReplacements As Boolean

        ''' <summary>
        ''' 
        ''' </summary>
        Public AskBeforeEachRealPhoneticFormReplacements As Boolean

        ''' <summary>
        ''' 
        ''' </summary>
        Public UseExampleWordSampling As Boolean

        ''' <summary>
        ''' 
        ''' </summary>
        Public MaximumNumberOfExampleWords As Integer

        ''' <summary>
        ''' 
        ''' </summary>
        Public OutputFolder As String

        ''' <summary>
        ''' This variable determines how many extra steps outside the original phonetic form the FST should go. Extra steps are needed if PDCS or PDS segments are added to the phonetic form.
        ''' </summary>
        Public UnresolvedJumpExtraSteps As Integer

        ''' <summary>
        ''' 
        ''' </summary>
        Public MaximumUnresolvedGraphemeJumps As Integer

        ''' <summary>
        ''' 
        ''' </summary>
        Public MaximumUnresolvedPhonemeJumps As Integer

        ''' <summary>
        ''' 
        ''' </summary>
        Public MaximumUnresolvedGraphemeAndPhonemeJumps As Integer

        ''' <summary>
        ''' 
        ''' </summary>
        Public AllowUnresolved_p2gs As Boolean

        ''' <summary>
        ''' 
        ''' </summary>
        Public ExportAttemptedSpellingSegmentations As Boolean

        ''' <summary>
        ''' 
        ''' </summary>
        Public ExportOnlyFailedAttemptedSpellingSegmentations As Boolean

        ''' <summary>
        ''' If set to true, no files are saved to disk.
        ''' </summary>
        Public DontExportAnything As Boolean = False

    Public Sub New(Optional InitializeDefaultSettings As Boolean = True, Optional p2g_rules_Path As String = "",
                       Optional UseTemporarySpellingChange As Boolean = True, Optional DontExportAnything As Boolean = False,
                       Optional IsRunFromServer As Boolean = False)

        Me.DontExportAnything = DontExportAnything
        Me.UseTemporarySpellingChange = UseTemporarySpellingChange
        If InitializeDefaultSettings = True Then SetDefaultValues(,, p2g_rules_Path, IsRunFromServer)

    End Sub

    ''' <summary>
    ''' Setting default values.
    ''' </summary>
    ''' <param name="WordStartMarker">A character to be used by the software to indicate word start. It can be any character which is NOT an orthographic character used in any of the orthographic forms analysed, nor a phonetic character used in the phonetic forms analysed. </param>
    ''' <param name="WordEndMarker">A character to be used by the software to indicate word end. It can be any character which is NOT an orthographic character used in any of the orthographic forms analysed, nor a phonetic character used in the phonetic forms analysed. </param>
    Public Sub SetDefaultValues(Optional ByVal WordStartMarker As String = "*",
                                    Optional ByVal WordEndMarker As String = "_",
                                    Optional p2g_rules_Path As String = "",
                                    Optional IsRunFromServer As Boolean = False,
                                Optional TemporarySpellingChangeFilePath As String = "")

        PhoneToGraphemesDictionary = New PhoneToGraphemes(WordStartMarker, WordEndMarker,, p2g_rules_Path)
        ErrorDictionary = Nothing
        ListOfGraphemes = Nothing
        'UseTemporarySpellingChange = True
        TemporarySpellingChangeListArray = Nothing
        WarnForIllegalCharactersInSpelling = False
        LengthSensitiveLookUpPhonemes = False
        UsePDCS_Addition = True
        UsePDCS_AdditionOnLastPhoneme = False
        UsePDCS_AdditionsWithoutLength = True
        UsePDS_Addition = True
        UsePDS_AdditionOnLastPhoneme = False
        UsePhonemeReplacement = False
        DoRealPhoneticFormReplacements = False
        AskBeforeEachRealPhoneticFormReplacements = False
        UseExampleWordSampling = True
        MaximumNumberOfExampleWords = 10
        OutputFolder = ""
        UnresolvedJumpExtraSteps = 5
        'MaximumUnresolvedGraphemeJumps = 4 ' These 3 are not set by default. Instead they are set for each word, depending on the word length.
        'MaximumUnresolvedPhonemeJumps = 4
        'MaximumUnresolvedGraphemeAndPhonemeJumps = 4
        AllowUnresolved_p2gs = True
        ExportAttemptedSpellingSegmentations = False
        ExportOnlyFailedAttemptedSpellingSegmentations = True

        'Loading temporary spelling changes
        If UseTemporarySpellingChange = True And IsRunFromServer = False And TemporarySpellingChangeFilePath <> "" Then
            SendInfoToLog("Loading a .txt-file containing Temporary Spelling Changes from " & TemporarySpellingChangeFilePath, , OutputFolder)
            TemporarySpellingChangeListArray = System.IO.File.ReadAllLines(TemporarySpellingChangeFilePath)
            SendInfoToLog("    Number of Temporary Spelling Change lines loaded " & TemporarySpellingChangeListArray.Count,, OutputFolder)
        End If

    End Sub

    Public Function CreateSettingsLogString() As String

        Dim Output As String = vbCrLf

        Output &= ”PhoneToGraphemesDictionary.WordStartMarker: ” & PhoneToGraphemesDictionary.WordStartMarker & vbCrLf
        Output &= ”PhoneToGraphemesDictionary.WordEndMarker: ” & PhoneToGraphemesDictionary.WordEndMarker & vbCrLf
        Output &= ”UseTemporarySpellingChange: ” & UseTemporarySpellingChange & vbCrLf
        If TemporarySpellingChangeListArray IsNot Nothing Then
            Output &= ”TemporarySpellingChangeListArray: ” & String.Join(vbCrLf, TemporarySpellingChangeListArray) & vbCrLf
        Else
            Output &= ”TemporarySpellingChangeListArray: ” & "No temporary spelling change list loaded." & vbCrLf
        End If
        Output &= ”WarnForIllegalCharactersInSpelling: ” & WarnForIllegalCharactersInSpelling & vbCrLf
        Output &= ”LengthSensitiveLookUpPhonemes: ” & LengthSensitiveLookUpPhonemes & vbCrLf
        Output &= ”UsePDCS_Addition: ” & UsePDCS_Addition & vbCrLf
        Output &= ”UsePDCS_AdditionOnLastPhoneme: ” & UsePDCS_AdditionOnLastPhoneme & vbCrLf
        Output &= ”UsePDCS_AdditionsWithoutLength: ” & UsePDCS_AdditionsWithoutLength & vbCrLf
        Output &= ”UsePDS_Addition: ” & UsePDS_Addition & vbCrLf
        Output &= ”UsePDS_AdditionOnLastPhoneme: ” & UsePDS_AdditionOnLastPhoneme & vbCrLf
        Output &= ”UsePhonemeReplacement: ” & UsePhonemeReplacement & vbCrLf
        Output &= ”DoRealPhoneticFormReplacements: ” & DoRealPhoneticFormReplacements & vbCrLf
        Output &= ”AskBeforeEachRealPhoneticFormReplacements: ” & AskBeforeEachRealPhoneticFormReplacements & vbCrLf
        Output &= ”UseExampleWordSampling: ” & UseExampleWordSampling & vbCrLf
        Output &= ”MaximumNumberOfExampleWords: ” & MaximumNumberOfExampleWords & vbCrLf
        Output &= ”OutputFolder: ” & OutputFolder & vbCrLf
        Output &= ”UnresolvedJumpExtraSteps: ” & UnresolvedJumpExtraSteps & vbCrLf
        Output &= ”AllowUnresolved_p2gs: ” & AllowUnresolved_p2gs & vbCrLf
        Output &= ”ExportAttemptedSpellingSegmentations: ” & ExportAttemptedSpellingSegmentations & vbCrLf
        Output &= ”ExportOnlyFailedAttemptedSpellingSegmentations: ” & ExportOnlyFailedAttemptedSpellingSegmentations & vbCrLf

        Return Output

    End Function

End Class



''' <summary>
''' This class calculates pronunciation-to-grapheme orthographic transparency.
''' </summary>
Public Class PhoneToGraphemes
    Inherits Dictionary(Of String, Phoneme)
    Property MaximumPhonemeCombinationLength As Integer
    Property ListOfPDSs As New List(Of String)
    ReadOnly Property WordStartMarker As String
    ReadOnly Property WordEndMarker As String
    Property NumberOfConcatenatedWordEndMarkers As Integer
    Property NormalizationCharacters As New List(Of String)
    ReadOnly Property Unresolved_p2g_Character As String
    ReadOnly Property UseReducedPhoneticForm As Boolean = False

    Public Class Phoneme
        Inherits List(Of Grapheme)
        Property PossibleDeletedCompundedSegments As Boolean = False
        Property PossibleDeletionSegment As Boolean = False
        Property PossibleDeletionSegmentSpellings As New List(Of String)
        Property PossibleReplacementPhonemes As New List(Of String)
        Property Comments As String = ""
        Property SimpleCount As Integer = 0
    End Class

    Public Class Grapheme
        Property PossibleSpelling As String
        Property PreAndPostPhonemeConditions As New List(Of String) '
        Property PreAndPostGraphemeConditions As New List(Of String) '
        Property Comments As String = ""
        Property SimpleCount As Integer = 0
        Property ExampleSamplingWithoutPDCSActive As Boolean = False 'TODO: Varför finns det ingen ExampleSamplingWithoutPDSActive, etc?
        Property ExampleSamplingWithPDCSActive As Boolean = False
        Property ExampleSamplingWithPDSActive As Boolean = False
        Property ExampleSamplingWithPhonemeReplacementActive As Boolean = False
        Property ExampleSamplingWithSilentgraphemesActive As Boolean = False
        Property ExamplesWithoutPDCS As New List(Of String)
        Property ExamplesWithPDCS As New List(Of String)
        Property ExamplesWithPDS As New List(Of String)
        Property ExamplesWithPhonemeReplacement As New List(Of String)
        Property ExamplesOfSilentgraphemes As New List(Of String)

        Property ExampleWordWithoutPDCS As Word
        Property ExampleWordWithPDCS As Word
        Property ExampleWordWithPDS As Word
        Property ExampleWordWithPhonemeReplacement As Word
        Property ExampleWordOfSilentgraphemes As Word

    End Class

    ''' <summary>
    ''' Sets the properties NonPDCSExampleSamplingActive and PDCSExampleSamplingActive to true in all graphemes of all phonemes in the current instance of PhoneToGraphemes.
    ''' </summary>
    Public Sub ActivateExampleSampling()
        For Each Phoneme In Me
            For Each Grapheme In Me(Phoneme.Key)
                Grapheme.ExampleSamplingWithoutPDCSActive = True
                Grapheme.ExampleSamplingWithPDCSActive = True
                Grapheme.ExampleSamplingWithPDSActive = True
                Grapheme.ExampleSamplingWithPhonemeReplacementActive = True
                Grapheme.ExampleSamplingWithSilentgraphemesActive = True
            Next
        Next
    End Sub


    Public Sub New(Optional ByRef SetWordStartMarker As String = "*", Optional ByRef SetWordEndMarker As String = "_",
                           Optional ByRef SetUnresolved_p2g_Character As String = "!",
                               Optional ByRef p2gRuleFilePath As String = "",
                           Optional SetUseReducedPhoneticForm As Boolean = False,
                       Optional ByVal TryLoadFrequencyData As Boolean = False)

        Unresolved_p2g_Character = SetUnresolved_p2g_Character
        WordStartMarker = SetWordStartMarker
        WordEndMarker = SetWordEndMarker
        UseReducedPhoneticForm = SetUseReducedPhoneticForm


        Dim inputData As String()

        If p2gRuleFilePath = "" Then
            Dim dataString As String = My.Resources.p2gRules
            dataString = dataString.Replace(vbCrLf, vbLf)
            inputData = dataString.Split(vbLf)
        Else
            inputData = File.ReadAllLines(p2gRuleFilePath)
        End If

        SendInfoToLog("Loading p2g keys from file " & p2gRuleFilePath)

        'Parsing input data (Should be structured: 
        '[phone]    (tab)   PDCS (optional) (tab)   comments
        'grapheme   (tab)   comma delimited PreContextConditionPhones    (tab) comma delimited PostContextConditionPhones   (tab) comma delimited PreAndPostGraphemeConditions(Structured: d-a-k, where the a is the actual phoneme. a is not used in the processing)   (tab)   comments

        Dim LastPhoneKey As String = ""

        For line = 1 To inputData.Length - 1 'First line contains headings

            Try
                'Skipping over empty lines, and lines with comments, starting with a ///
                If inputData(line).Trim = "" Or inputData(line).TrimStart.StartsWith("///") Then
                    'Continues to next line
                    Continue For

                Else

                    Dim lineTabSplit() As String = inputData(line).Split(vbTab)

                    'Checking for normalization characters line:
                    If lineTabSplit(0).TrimStart.StartsWith("NormChars") Then
                        If lineTabSplit.Length > 1 Then
                            Dim NormChars() As String = lineTabSplit(1).Split(" ")
                            For n = 0 To NormChars.Length - 1
                                If NormChars(n).Trim <> "" Then
                                    Me.NormalizationCharacters.Add(NormChars(n).Trim)
                                End If
                            Next
                        End If
                        Continue For
                    End If


                    'Checking if its a new-phone line
                    If lineTabSplit(0).StartsWith("[") Then
                        'Its a new phoneme, adds it
                        LastPhoneKey = lineTabSplit(0).TrimStart("[").TrimEnd("]")
                        Me.Add(LastPhoneKey, New PhoneToGraphemes.Phoneme)

                        If lineTabSplit.Length > 1 Then
                            Try
                                If lineTabSplit(1).Trim = "PDCS" Then Me(LastPhoneKey).PossibleDeletedCompundedSegments = True
                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                        End If

                        'Adding possible deletion segment setting, and spellings
                        If lineTabSplit.Length > 2 Then
                            Try
                                If lineTabSplit(2).Trim = "PDS" Then Me(LastPhoneKey).PossibleDeletionSegment = True
                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                        End If

                        'Adding possible replacement segment setting, and spellings
                        If lineTabSplit.Length > 3 Then
                            Try
                                If lineTabSplit(3).Trim <> "" Then
                                    Dim ReplacementPhonemes() As String = lineTabSplit(3).Trim.TrimStart("[").TrimEnd("]").Split(", ")

                                    For n = 0 To ReplacementPhonemes.Length - 1
                                        If ReplacementPhonemes(n).Trim <> "" Then
                                            'Adding the replacement phonemes
                                            Me(LastPhoneKey).PossibleReplacementPhonemes.Add(ReplacementPhonemes(n).Trim)

                                        Else
                                            Throw New Exception("The replacement phoneme specification For the phoneme [" & LastPhoneKey & "] Is incorrect." & vbCrLf &
                                                                        "An example Of the format you should use (put it In the fourth phoneme column)" & vbCrLf &
                                                                        "[t,pː]" & vbCrLf &
                                                                        "Here [t] And [pː] are alternative replacement phonemes For the current phoneme")
                                        End If
                                    Next
                                End If


                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                        End If

                        'Adding phoneme comments
                        If lineTabSplit.Length > 4 Then
                            Try
                                Me(LastPhoneKey).Comments = lineTabSplit(4).Trim
                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                        End If

                        If TryLoadFrequencyData = True Then
                            'Adding phoneme frequency data if available
                            If lineTabSplit.Length > 5 Then
                                Try
                                    If IsNumeric(lineTabSplit(5).Trim) Then Me(LastPhoneKey).SimpleCount = lineTabSplit(5).Trim
                                Catch ex As Exception
                                    MsgBox(ex.ToString)
                                End Try
                            End If
                        End If

                    Else
                        'Its a grapheme data line (also adds empty lines to possible spellings! Important to not have any extra empty lines in the input file!)

                        'Creates a new grapheme, and adds its content
                        Dim NewGrapheme As New PhoneToGraphemes.Grapheme
                        NewGrapheme.PossibleSpelling = lineTabSplit(0).Trim


                        'Adds pre and post grapheme phoneme conditions
                        If lineTabSplit.Length > 1 Then
                            If lineTabSplit(1).Trim <> "" Then
                                Dim PreAndPostPhonemeContextSplit() As String = lineTabSplit(1).Split(",")
                                For n = 0 To PreAndPostPhonemeContextSplit.Length - 1
                                    If Not PreAndPostPhonemeContextSplit(n).Trim = "" Then
                                        'Adding tab delimited pre and post conditions (not the current grapheme!)
                                        Try
                                            NewGrapheme.PreAndPostPhonemeConditions.Add(PreAndPostPhonemeContextSplit(n).Trim.Split("-")(0).Trim & vbTab & PreAndPostPhonemeContextSplit(n).Trim.Split("-")(2).Trim)
                                        Catch ex As Exception
                                            MsgBox("Error adding pre-And post grapheme phoneme condition " & lineTabSplit(1) & " of phoneme: " & LastPhoneKey)
                                        End Try
                                    End If
                                Next
                            End If
                        End If

                        'Adds pre and post grapheme spelling conditions
                        If lineTabSplit.Length > 2 Then
                            If lineTabSplit(2).Trim <> "" Then
                                Dim PreAndPostGraphemeContextSplit() As String = lineTabSplit(2).Split(",")
                                For n = 0 To PreAndPostGraphemeContextSplit.Length - 1
                                    If Not PreAndPostGraphemeContextSplit(n).Trim = "" Then
                                        'Adding tab delimited pre and post conditions (not the current grapheme!)
                                        Try
                                            NewGrapheme.PreAndPostGraphemeConditions.Add(PreAndPostGraphemeContextSplit(n).Trim.Split("-")(0).Trim & vbTab & PreAndPostGraphemeContextSplit(n).Trim.Split("-")(2).Trim)
                                        Catch ex As Exception
                                            MsgBox("Error adding pre-And post grapheme condition " & lineTabSplit(2) & " of phoneme: " & LastPhoneKey)
                                        End Try
                                    End If
                                Next
                            End If
                        End If

                        '
                        If lineTabSplit.Length > 3 Then
                            Try
                                Dim Column3 As String = lineTabSplit(3).Trim

                                If Column3.Trim <> "" Then
                                    MsgBox("The fourth grapheme column should be blank. Any information written here will be discarded." & vbCr & vbCr &
                                               "The fourth grapheme column of phoneme [" & LastPhoneKey & "] And grapheme <" & NewGrapheme.PossibleSpelling &
                                               "> contains the following data " & vbCr & vbCr &
                                               Column3 & vbCr & vbCr & "Press Ok to discard this data And contiune.")
                                End If

                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                        End If

                        'Adding any comments
                        If lineTabSplit.Length > 4 Then
                            Try
                                NewGrapheme.Comments = lineTabSplit(4).Trim
                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                        End If

                        If TryLoadFrequencyData = True Then
                            'Adding any frequency data
                            If lineTabSplit.Length > 5 Then
                                Try
                                    If IsNumeric(lineTabSplit(5).Trim) Then NewGrapheme.SimpleCount = lineTabSplit(5).Trim
                                Catch ex As Exception
                                    MsgBox(ex.ToString)
                                End Try
                            End If

                        End If


                        'Adds the new grapheme to the last encounterred phone
                        Me(LastPhoneKey).Add(NewGrapheme)
                    End If

                End If

            Catch ex As Exception
                MsgBox("Error reading line " & inputData(line) & vbCrLf & vbCrLf & ex.ToString)
            End Try

        Next

        'Adds word start character
        Me.Add(WordStartMarker, New PhoneToGraphemes.Phoneme)
        Dim WordStartGrapheme As New PhoneToGraphemes.Grapheme
        WordStartGrapheme.PossibleSpelling = WordStartMarker
        Me(WordStartMarker).Add(WordStartGrapheme)

        If WordStartMarker <> WordEndMarker Then
            'Adding also the word end character, if its different from the word start marker
            Me.Add(WordEndMarker, New PhoneToGraphemes.Phoneme)
            Dim WordEndGrapheme As New PhoneToGraphemes.Grapheme
            WordEndGrapheme.PossibleSpelling = WordEndMarker
            Me(WordEndMarker).Add(WordEndGrapheme)
        End If

        'Sets the class word start and word end markers
        Me.WordStartMarker = WordStartMarker
        Me.WordEndMarker = WordEndMarker

        'Output.Add(WordEndMarker & " " & WordEndMarker, New PhoneToGraphemes.Phoneme)
        'Dim WordBournaryGrapheme2 As New PhoneToGraphemes.Grapheme
        'WordBournaryGrapheme2.PossibleSpelling = WordEndMarker
        'Output(WordEndMarker & " " & WordEndMarker).Add(WordBournaryGrapheme2)


        'Determining the length of the longest phoneme combinations
        Me.MaximumPhonemeCombinationLength = 0
        For Each PhonemeCombination In Me
            Dim CurrentPhonemeCombinationLength As Integer = PhonemeCombination.Key.Split(" ").Length
            If CurrentPhonemeCombinationLength > Me.MaximumPhonemeCombinationLength Then
                Me.MaximumPhonemeCombinationLength = CurrentPhonemeCombinationLength
            End If
        Next

        'Summing PDSs
        For Each Phoneme In Me
            If Phoneme.Value.PossibleDeletionSegment = True Then Me.ListOfPDSs.Add(Phoneme.Key)
        Next

        If Me.NormalizationCharacters.Count = 0 Then MsgBox("Please note that no normalization characters were loaded to the p2g rule file!")

        SendInfoToLog(" p2g keys were loaded successfully.")

    End Sub

    Public Sub Export_p2g_Examples(ByVal saveDirectory As String, Optional ByRef saveFileName As String = "",
                                                Optional BoxTitle As String = "Choose location to store the p2g output file...")

        Try

            SendInfoToLog("Attempting to export p2g dictionary with example words")

            'Choosing file location
            Dim filepath As String = Path.Combine(saveDirectory, saveFileName & ".txt")
            If Not Directory.Exists(Path.GetDirectoryName(filepath)) Then Directory.CreateDirectory(Path.GetDirectoryName(filepath))

            'Declaring counters
            Dim NumberOfPhonemes As Integer = 0
            Dim NumberOfGraphemes As Integer = 0
            Dim NumberOfNonPDCS_GraphemesWithoutExamples As Integer = 0
            Dim NumberOfPDCS_GraphemesWithoutExamples As Integer = 0

            'Save it to file
            Dim writer As New StreamWriter(filepath, False, Text.Encoding.UTF8)

            writer.WriteLine("Phon/Graph" & vbTab & "PDCS/PhoneticContextConditions" & vbTab & "PDS/SpellingContextConditions" & vbTab & "ReplPhonemes" & vbTab &
                                         "Comments" & vbTab & "GraphemeCount" & vbTab & "ExampleWords_NoPDCS" & vbTab & "ExampleWords_PDCS" & vbTab & "ExampleWords_PDS" & vbTab & "ExampleWords_SilentGraphemes")

            writer.WriteLine(vbCrLf & "NormChars" & vbTab & String.Join(" ", NormalizationCharacters) & vbCrLf & vbCrLf)

            For Each Phoneme In Me

                'Skips to next if it's the word start or word end markers
                If Phoneme.Key = WordStartMarker Or Phoneme.Key = WordEndMarker Then Continue For

                NumberOfPhonemes += 1

                Dim PDCS_String As String = ""
                If Phoneme.Value.PossibleDeletedCompundedSegments = True Then PDCS_String = "PDCS"
                Dim PDS_String As String = ""
                If Phoneme.Value.PossibleDeletionSegment = True Then PDS_String = "PDS"
                Dim PR_String As String = ""
                If Phoneme.Value.PossibleReplacementPhonemes.Count > 0 Then PR_String = "[" & String.Join(",", Phoneme.Value.PossibleReplacementPhonemes) & "]"

                writer.WriteLine("[" & Phoneme.Key & "]" & vbTab & PDCS_String & vbTab & PDS_String & vbTab & PR_String & vbTab & Phoneme.Value.Comments & vbTab & Phoneme.Value.SimpleCount)

                For Each Grapheme In Me(Phoneme.Key)
                    NumberOfGraphemes += 1

                    Dim PreAndPostPhonemeContextConditions As New List(Of String)
                    For n = 0 To Grapheme.PreAndPostPhonemeConditions.Count - 1
                        Dim CurrentSplit() As String = Grapheme.PreAndPostPhonemeConditions(n).Split(vbTab)
                        PreAndPostPhonemeContextConditions.Add(CurrentSplit(0) & "-" & Phoneme.Key & "-" & CurrentSplit(1))
                    Next
                    Dim PreAndPostPhonemeContextConditionsString As String = String.Join(", ", PreAndPostPhonemeContextConditions)

                    Dim PreAndPostGraphemeContextConditions As New List(Of String)
                    For n = 0 To Grapheme.PreAndPostGraphemeConditions.Count - 1
                        Dim CurrentSplit() As String = Grapheme.PreAndPostGraphemeConditions(n).Split(vbTab)
                        PreAndPostGraphemeContextConditions.Add(CurrentSplit(0) & "-" & Grapheme.PossibleSpelling & "-" & CurrentSplit(1))
                    Next
                    Dim PreAndPostGraphemeContextConditionsString As String = String.Join(", ", PreAndPostGraphemeContextConditions)

                    Dim ExampleWords_NoPDCS As String = String.Join(", ", Grapheme.ExamplesWithoutPDCS)
                    If Grapheme.ExamplesWithoutPDCS.Count = 0 Then NumberOfNonPDCS_GraphemesWithoutExamples += 1

                    Dim ExampleWords_PDCS As String = String.Join(", ", Grapheme.ExamplesWithPDCS)
                    If Grapheme.ExamplesWithPDCS.Count = 0 Then NumberOfPDCS_GraphemesWithoutExamples += 1

                    Dim ExampleWords_PDS As String = String.Join(", ", Grapheme.ExamplesWithPDS)

                    Dim ExampleWords_PR As String = String.Join(", ", Grapheme.ExamplesWithPhonemeReplacement)

                    Dim ExampleWords_SilentGraphemes As String = String.Join(", ", Grapheme.ExamplesOfSilentgraphemes)

                    writer.WriteLine(Grapheme.PossibleSpelling & vbTab &
                                                 PreAndPostPhonemeContextConditionsString & vbTab &
                                                 PreAndPostGraphemeContextConditionsString & vbTab &
                                                 vbTab & 'Empty column to align with the phoneme comments
                                                 Grapheme.Comments & vbTab &
                                                 Grapheme.SimpleCount & vbTab &
                                                 ExampleWords_NoPDCS.Replace(WordStartMarker, "").Replace(WordEndMarker, "") & vbTab &
                                                 ExampleWords_PDCS.Replace(WordStartMarker, "").Replace(WordEndMarker, "") & vbTab &
                                                 ExampleWords_PDS.Replace(WordStartMarker, "").Replace(WordEndMarker, "") & vbTab &
                                                 ExampleWords_PR.Replace(WordStartMarker, "").Replace(WordEndMarker, "") & vbTab &
                                                 ExampleWords_SilentGraphemes.Replace(WordStartMarker, "").Replace(WordEndMarker, ""))

                Next
                writer.WriteLine()
            Next

            writer.Close()

            SendInfoToLog("     The p2g dictionary file with example words has been saved to: " & filepath & vbCrLf &
                                      "     Results:" & vbCrLf &
                                      "     Number of phonemes: " & NumberOfPhonemes & vbCrLf &
                                      "     Total number of graphemes (for all phonemes): " & NumberOfGraphemes & vbCrLf &
                                      "     Number of non PDCS graphemes without example words: " & NumberOfNonPDCS_GraphemesWithoutExamples & vbCrLf &
                                      "     Number of PDCS graphemes without example words: " & NumberOfPDCS_GraphemesWithoutExamples)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub

    Public Sub Export_p2g_MostCommonExample(ByVal saveDirectory As String, Optional ByRef saveFileName As String = "",
                                                Optional BoxTitle As String = "Choose location to store the p2g output file containing the most common examples...")

        Try

            SendInfoToLog("Attempting to export p2g dictionary with example words")

            'Choosing file location
            Dim filepath As String = Path.Combine(saveDirectory, saveFileName & ".txt")
            If Not Directory.Exists(Path.GetDirectoryName(filepath)) Then Directory.CreateDirectory(Path.GetDirectoryName(filepath))

            'Declaring counters
            Dim NumberOfPhonemes As Integer = 0
            Dim NumberOfGraphemes As Integer = 0
            Dim NumberOfNonPDCS_GraphemesWithoutExamples As Integer = 0
            Dim NumberOfPDCS_GraphemesWithoutExamples As Integer = 0

            'Save it to file
            Dim writer As New StreamWriter(filepath, False, Text.Encoding.UTF8)

            writer.WriteLine("Phon/Graph" & vbTab & "PDCS/PhoneticContextConditions" & vbTab & "PDS/SpellingContextConditions" & vbTab & "ReplPhonemes" & vbTab &
                                         "Comments" & vbTab & "GraphemeCount" & vbTab & "ExampleWords_NoPDCS" & vbTab & "ExampleWords_PDCS" & vbTab & "ExampleWords_PDS" & vbTab & "ExampleWords_SilentGraphemes")

            writer.WriteLine(vbCrLf & "NormChars" & vbTab & String.Join(" ", NormalizationCharacters) & vbCrLf & vbCrLf)

            For Each Phoneme In Me

                'Skips to next if it's the word start or word end markers
                If Phoneme.Key = WordStartMarker Or Phoneme.Key = WordEndMarker Then Continue For

                NumberOfPhonemes += 1

                Dim PDCS_String As String = ""
                If Phoneme.Value.PossibleDeletedCompundedSegments = True Then PDCS_String = "PDCS"
                Dim PDS_String As String = ""
                If Phoneme.Value.PossibleDeletionSegment = True Then PDS_String = "PDS"
                Dim PR_String As String = ""
                If Phoneme.Value.PossibleReplacementPhonemes.Count > 0 Then PR_String = "[" & String.Join(",", Phoneme.Value.PossibleReplacementPhonemes) & "]"

                writer.WriteLine("[" & Phoneme.Key & "]" & vbTab & PDCS_String & vbTab & PDS_String & vbTab & PR_String & vbTab & Phoneme.Value.Comments & vbTab & Phoneme.Value.SimpleCount)

                For Each Grapheme In Me(Phoneme.Key)
                    NumberOfGraphemes += 1

                    Dim PreAndPostPhonemeContextConditions As New List(Of String)
                    For n = 0 To Grapheme.PreAndPostPhonemeConditions.Count - 1
                        Dim CurrentSplit() As String = Grapheme.PreAndPostPhonemeConditions(n).Split(vbTab)
                        PreAndPostPhonemeContextConditions.Add(CurrentSplit(0) & "-" & Phoneme.Key & "-" & CurrentSplit(1))
                    Next
                    Dim PreAndPostPhonemeContextConditionsString As String = String.Join(", ", PreAndPostPhonemeContextConditions)

                    Dim PreAndPostGraphemeContextConditions As New List(Of String)
                    For n = 0 To Grapheme.PreAndPostGraphemeConditions.Count - 1
                        Dim CurrentSplit() As String = Grapheme.PreAndPostGraphemeConditions(n).Split(vbTab)
                        PreAndPostGraphemeContextConditions.Add(CurrentSplit(0) & "-" & Grapheme.PossibleSpelling & "-" & CurrentSplit(1))
                    Next
                    Dim PreAndPostGraphemeContextConditionsString As String = String.Join(", ", PreAndPostGraphemeContextConditions)

                    'ExampleWordWithoutPDCS
                    Dim ExampleWords_NoPDCS As String = ""
                    If Not Grapheme.ExampleWordWithoutPDCS Is Nothing Then
                        If UseReducedPhoneticForm = True Then
                            ExampleWords_NoPDCS = "<" & Grapheme.ExampleWordWithoutPDCS.OrthographicForm & ">, [" & String.Concat(Grapheme.ExampleWordWithoutPDCS.BuildReducedIpaArray()) & "]"
                        Else
                            ExampleWords_NoPDCS = "<" & Grapheme.ExampleWordWithoutPDCS.OrthographicForm & ">, [" & String.Concat(Grapheme.ExampleWordWithoutPDCS.BuildExtendedIpaArray(,,,,, False, False)) & "]"
                        End If
                    Else
                        NumberOfNonPDCS_GraphemesWithoutExamples += 1
                    End If

                    'ExamplesWithPDCS
                    Dim ExampleWords_PDCS As String = ""
                    If Not Grapheme.ExampleWordWithPDCS Is Nothing Then
                        If UseReducedPhoneticForm = True Then
                            ExampleWords_PDCS = "<" & Grapheme.ExampleWordWithPDCS.OrthographicForm & ">, [" & String.Concat(Grapheme.ExampleWordWithPDCS.BuildReducedIpaArray()) & "]"
                        Else
                            ExampleWords_PDCS = "<" & Grapheme.ExampleWordWithPDCS.OrthographicForm & ">, [" & String.Concat(Grapheme.ExampleWordWithPDCS.BuildExtendedIpaArray(,,,,, False, False)) & "]"
                        End If
                    Else
                        NumberOfPDCS_GraphemesWithoutExamples += 1
                    End If


                    'ExampleWords_PDS
                    Dim ExampleWords_PDS As String = ""
                    If Not Grapheme.ExampleWordWithPDS Is Nothing Then
                        If UseReducedPhoneticForm = True Then
                            ExampleWords_PDS = "<" & Grapheme.ExampleWordWithPDS.OrthographicForm & ">, [" & String.Concat(Grapheme.ExampleWordWithPDS.BuildReducedIpaArray()) & "]"
                        Else
                            ExampleWords_PDS = "<" & Grapheme.ExampleWordWithPDS.OrthographicForm & ">, [" & String.Concat(Grapheme.ExampleWordWithPDS.BuildExtendedIpaArray(,,,,, False, False)) & "]"
                        End If
                    Else
                        'NumberOfPDCS_GraphemesWithoutExamples += 1
                    End If


                    'ExampleWordWithPhonemeReplacement
                    Dim ExampleWords_PR As String = ""
                    If Not Grapheme.ExampleWordWithPhonemeReplacement Is Nothing Then
                        If UseReducedPhoneticForm = True Then
                            ExampleWords_PR = "<" & Grapheme.ExampleWordWithPhonemeReplacement.OrthographicForm & ">, [" & String.Concat(Grapheme.ExampleWordWithPhonemeReplacement.BuildReducedIpaArray()) & "]"
                        Else
                            ExampleWords_PR = "<" & Grapheme.ExampleWordWithPhonemeReplacement.OrthographicForm & ">, [" & String.Concat(Grapheme.ExampleWordWithPhonemeReplacement.BuildExtendedIpaArray(,,,,, False, False)) & "]"
                        End If
                    Else
                        'NumberOfPDCS_GraphemesWithoutExamples += 1
                    End If

                    'ExampleWordOfSilentgraphemes
                    Dim ExampleWords_SilentGraphemes As String = ""
                    If Not Grapheme.ExampleWordOfSilentgraphemes Is Nothing Then
                        If UseReducedPhoneticForm = True Then
                            ExampleWords_SilentGraphemes = "<" & Grapheme.ExampleWordOfSilentgraphemes.OrthographicForm & ">, [" & String.Concat(Grapheme.ExampleWordOfSilentgraphemes.BuildReducedIpaArray()) & "]"
                        Else
                            ExampleWords_SilentGraphemes = "<" & Grapheme.ExampleWordOfSilentgraphemes.OrthographicForm & ">, [" & String.Concat(Grapheme.ExampleWordOfSilentgraphemes.BuildExtendedIpaArray(,,,,, False, False)) & "]"
                        End If
                    Else
                        'NumberOfPDCS_GraphemesWithoutExamples += 1
                    End If



                    writer.WriteLine(Grapheme.PossibleSpelling & vbTab &
                                                 PreAndPostPhonemeContextConditionsString & vbTab &
                                                 PreAndPostGraphemeContextConditionsString & vbTab &
                                                 vbTab & 'Empty column to align with the phoneme comments
                                                 Grapheme.Comments & vbTab &
                                                 Grapheme.SimpleCount & vbTab &
                                                 ExampleWords_NoPDCS & vbTab &
                                                 ExampleWords_PDCS & vbTab &
                                                 ExampleWords_PDS & vbTab &
                                                 ExampleWords_PR & vbTab &
                                                 ExampleWords_SilentGraphemes)

                Next
                writer.WriteLine()
            Next

            writer.Close()

            SendInfoToLog("     The p2g dictionary file with example words has been saved to: " & filepath & vbCrLf &
                                      "     Results:" & vbCrLf &
                                      "     Number of phonemes: " & NumberOfPhonemes & vbCrLf &
                                      "     Total number of graphemes (for all phonemes): " & NumberOfGraphemes & vbCrLf &
                                      "     Number of non PDCS graphemes without example words: " & NumberOfNonPDCS_GraphemesWithoutExamples & vbCrLf &
                                      "     Number of PDCS graphemes without example words: " & NumberOfPDCS_GraphemesWithoutExamples)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub


    Public Function CreateDictionaryOfMostCommonP2G() As Dictionary(Of String, String)

        Dim Output As New Dictionary(Of String, String)

        For Each Phoneme In Me

            Dim MostCommonGrapheme As String = ""
            Dim MostCommonGraphemeFrequency As Integer = -1 'Should trigger storage on the first run
            For Each Grapheme In Me(Phoneme.Key)

                'Stores the most common grapheme for the current phoneme
                If Grapheme.SimpleCount > MostCommonGraphemeFrequency Then
                    MostCommonGrapheme = Grapheme.PossibleSpelling
                    MostCommonGraphemeFrequency = Grapheme.SimpleCount
                End If
            Next

            Output.Add(Phoneme.Key, MostCommonGrapheme)
        Next

        Return Output

    End Function

End Class

''' <summary>
''' This class calculates the type of Grapheme-to-pronunciation orthographic transparency (G2P-OT) described by Berndt, Reggia, and Mitchum (1987).
''' </summary>
Public Class GraphemeToPhonemes
    Inherits Dictionary(Of String, Grapheme)
    Property Total_FrequencyData As Double 'TW for token Weighted
    Property WordStartMarker As String
    Property WordEndMarker As String
    Property Unresolved_p2g_Character As String
    Property CalculationUnit As WordGroup.WordFrequencyUnit = WordGroup.WordFrequencyUnit.WordType

    Public Class Grapheme
        Inherits Dictionary(Of String, PhonemeData)
        Property FrequencyData As Double
        Property PriorProbability As Double = 0
        Property HighestConditionalProbability As Double = 0

    End Class

    Public Class PhonemeData
        Property FrequencyData As Double
        Property ConditionalProbability As Double = 0
        Property ConditionalPredictability As Double = 0
        Property ExampleWord As Word
    End Class

    Private Class StringGraphemeCombination
        Property MyString As String
        Property MyGrapheme As GraphemeToPhonemes.Grapheme
    End Class


    Public Sub New(Optional ByRef SetWordStartMarker As String = "*", Optional ByRef SetWordEndMarker As String = "_",
                                       Optional ByVal SetUnresolved_p2g_Character As String = "!",
                           Optional ByRef SetCalculationUnit As WordGroup.WordFrequencyUnit = WordGroup.WordFrequencyUnit.WordType)

        WordStartMarker = SetWordStartMarker
        WordEndMarker = SetWordEndMarker
        Unresolved_p2g_Character = SetUnresolved_p2g_Character
        CalculationUnit = SetCalculationUnit

    End Sub



    Public Sub AddGraphemeData(ByRef CurrentWord As Word, ByRef GraphemeKey As String,
                                       ByRef CurrentPhonemeString As String, Optional ByRef ExampleWordFrequencyLimit As Integer = 10,
                                       Optional ByRef DoNotAddDataWithUnresolved_p2g As Boolean = True)

        If DoNotAddDataWithUnresolved_p2g = True Then
            If GraphemeKey.Contains(Unresolved_p2g_Character) Or CurrentPhonemeString.Contains(Unresolved_p2g_Character) Then Exit Sub
        End If

        Dim FrequencyData As Double = 0
        Select Case CalculationUnit
            Case WordGroup.WordFrequencyUnit.RawFrequency
                FrequencyData = CurrentWord.RawWordTypeFrequency
            Case WordGroup.WordFrequencyUnit.WordType
                FrequencyData = 1
            Case Else
                Throw New NotImplementedException
        End Select

        'Checking if the FirstGraphemeInBlock exists
        If Not Me.ContainsKey(GraphemeKey) Then

            'Adds the grapheme
            Dim newGraphemeBlockData As New GraphemeToPhonemes.Grapheme
            Me.Add(GraphemeKey, newGraphemeBlockData)
            Me(GraphemeKey).FrequencyData = FrequencyData

            'Adds phoneme data
            Dim newPhonemeData As New GraphemeToPhonemes.PhonemeData
            Me(GraphemeKey).Add(CurrentPhonemeString, newPhonemeData)
            Me(GraphemeKey)(CurrentPhonemeString).FrequencyData = FrequencyData

            'Not adding Hapax legomenas
            If Not CurrentWord.RawWordTypeFrequency >= ExampleWordFrequencyLimit Then

                'Not adding example words if they contain an unresolved character (as this may acctually cause erroneous segmentation)
                If Not String.Concat(CurrentWord.Sonographs_Letters).Contains(Unresolved_p2g_Character) Then
                    Me(GraphemeKey)(CurrentPhonemeString).ExampleWord = CurrentWord
                End If
            End If

        Else
            'Increases GraphemeBlockKey count
            Me(GraphemeKey).FrequencyData += FrequencyData

            If Not Me(GraphemeKey).ContainsKey(CurrentPhonemeString) Then

                'Adds phoneme data
                Dim newPhonemeData As New GraphemeToPhonemes.PhonemeData
                Me(GraphemeKey).Add(CurrentPhonemeString, newPhonemeData)
                Me(GraphemeKey)(CurrentPhonemeString).FrequencyData = FrequencyData

                'Not adding Hapax legomenas
                If Not CurrentWord.RawWordTypeFrequency >= ExampleWordFrequencyLimit Then

                    'Not adding example words if they contain an unresolved character (as this may acctually cause erroneous segmentation)
                    If Not String.Concat(CurrentWord.Sonographs_Letters).Contains(Unresolved_p2g_Character) Then
                        Me(GraphemeKey)(CurrentPhonemeString).ExampleWord = CurrentWord
                    End If
                End If

            Else

                'Increases phoneme count
                Me(GraphemeKey)(CurrentPhonemeString).FrequencyData += FrequencyData

                'Example word selection
                'Not adding Hapax legomenas
                If Not CurrentWord.RawWordTypeFrequency >= ExampleWordFrequencyLimit Then

                    'Not adding example words if they contain an unresolved character (as this may acctually cause erroneous segmentation)
                    If Not String.Concat(CurrentWord.Sonographs_Letters).Contains(Unresolved_p2g_Character) Then

                        'Adding the word if no word has been previously added
                        If Me(GraphemeKey)(CurrentPhonemeString).ExampleWord Is Nothing Then
                            Me(GraphemeKey)(CurrentPhonemeString).ExampleWord = CurrentWord
                        Else

                            'Replacing the existing word if it is marked as foreign, and the new one is not
                            If Me(GraphemeKey)(CurrentPhonemeString).ExampleWord.ForeignWord = True And CurrentWord.ForeignWord = False Then
                                Me(GraphemeKey)(CurrentPhonemeString).ExampleWord = CurrentWord
                            Else

                                'Replacing the existing word if contains a special character, and new one is not
                                If Me(GraphemeKey)(CurrentPhonemeString).ExampleWord.OrthographicFormContainsSpecialCharacter = True And CurrentWord.OrthographicFormContainsSpecialCharacter = False Then
                                    Me(GraphemeKey)(CurrentPhonemeString).ExampleWord = CurrentWord
                                Else

                                    'Replaces the existing word if it has a phoneme range is 3-8 phonemes and a higher phonotactic probability (for publication purpose)
                                    Dim PhonemeCount As Integer = CurrentWord.CountPhonemes
                                    If PhonemeCount > 2 And PhonemeCount < 9 Then

                                        Dim NewWordValues As New List(Of Double) From {CurrentWord.SSPP_Average, CurrentWord.RawWordTypeFrequency}
                                        Dim ExistingWordValues As New List(Of Double) From {Me(GraphemeKey)(CurrentPhonemeString).ExampleWord.SSPP_Average,
                                                    Me(GraphemeKey)(CurrentPhonemeString).ExampleWord.RawWordTypeFrequency}

                                        If GeometricMean(NewWordValues) > GeometricMean(ExistingWordValues) Then
                                            Me(GraphemeKey)(CurrentPhonemeString).ExampleWord = CurrentWord
                                        End If

                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub



    ''' <summary>
    ''' Using the collected g2p data, probability data for each grapheme to phoneme correspondence is calculated (in accordance with Berndt, R. S., Reggia, J. A., and Mitchum, C. C. (1987))
    ''' </summary>
    Public Sub CalculateProbabilityData()

        Total_FrequencyData = 0

        For Each grapheme As KeyValuePair(Of String, Grapheme) In Me

            'Summing total grapheme frequency
            grapheme.Value.FrequencyData = 0
            For Each phoneme In grapheme.Value
                grapheme.Value.FrequencyData += phoneme.Value.FrequencyData
            Next

            'Calculating conditional grapheme to phoneme probabilities
            For Each phoneme In grapheme.Value
                phoneme.Value.ConditionalProbability = phoneme.Value.FrequencyData / grapheme.Value.FrequencyData
            Next

            'Summing total grapheme counts (for calculation of prior probability)
            Total_FrequencyData += grapheme.Value.FrequencyData

        Next

        'Determining the highest g2p probability for each current grapheme (I.e. the probability of the most frequent correspondence of a grapheme)
        For Each grapheme As KeyValuePair(Of String, Grapheme) In Me
            'Resetting HighestConditionalProbability
            grapheme.Value.HighestConditionalProbability = 0

            'Going through each phoneme and determining storing the highest conditional probability in grapheme.Value.HighestConditionalProbability
            For Each phoneme In grapheme.Value
                If phoneme.Value.ConditionalProbability > grapheme.Value.HighestConditionalProbability Then
                    grapheme.Value.HighestConditionalProbability = phoneme.Value.ConditionalProbability
                End If
            Next
        Next

        'Calculating prior probabilities
        For Each grapheme As KeyValuePair(Of String, Grapheme) In Me
            grapheme.Value.PriorProbability = grapheme.Value.FrequencyData / Total_FrequencyData
        Next


        'Calculating g2p predictability
        For Each CurrentGrapheme As KeyValuePair(Of String, Grapheme) In Me
            For Each CurrentPhoneme In CurrentGrapheme.Value
                CurrentPhoneme.Value.ConditionalPredictability = CurrentPhoneme.Value.ConditionalProbability / CurrentGrapheme.Value.HighestConditionalProbability
            Next
        Next


        'H. Setting the Predictabilities and probabilties for unresolved p2gs to 0
        If Me.ContainsKey(Unresolved_p2g_Character) Then
            Me(Unresolved_p2g_Character).HighestConditionalProbability = 0

            For Each CurrentPhoneme In Me(Unresolved_p2g_Character)
                CurrentPhoneme.Value.ConditionalProbability = 0
                CurrentPhoneme.Value.ConditionalPredictability = 0
            Next
        End If

        SendInfoToLog("Finished calculating spelling probabilities.")

    End Sub

    Public Sub Collect_G2P_Data(ByRef InputWordList As WordGroup)

        'Going through all member words and collect frequency data, and stores it in fb2gData
        For word = 0 To InputWordList.MemberWords.Count - 1

            Dim TempSpelling As String = WordStartMarker & InputWordList.MemberWords(word).OrthographicForm & WordEndMarker
            Dim OrthStartIndex As Integer = 0

            For CurrentGraphemeBlockIndex = 0 To InputWordList.MemberWords(word).Sonographs_Letters.Count - 1

                'Adding gb2p Data
                Try

                    AddGraphemeData(InputWordList.MemberWords(word), InputWordList.MemberWords(word).Sonographs_Letters(CurrentGraphemeBlockIndex),
                                                        InputWordList.MemberWords(word).Sonographs_Pronunciation(CurrentGraphemeBlockIndex))

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try


                'Increasing OrthStartIndex 
                OrthStartIndex += InputWordList.MemberWords(word).Sonographs_Letters(CurrentGraphemeBlockIndex).Length

            Next
        Next

        'Calculates spelling probability
        CalculateProbabilityData()

    End Sub


    Public Sub Apply_G2P_Data(ByRef InputWordList As WordGroup)

        'Going through all member words and looks up their spelling probability values
        For word = 0 To InputWordList.MemberWords.Count - 1

            'Resetting spellingRegularity
            InputWordList.MemberWords(word).G2P_OT.Clear()


            Dim TempSpelling As String = WordStartMarker & InputWordList.MemberWords(word).OrthographicForm & WordEndMarker
            Dim OrthStartIndex As Integer = 0

            For CurrentGraphemeBlockIndex = 0 To InputWordList.MemberWords(word).Sonographs_Letters.Count - 1

                'Calculating spelling regularity data
                Dim CurrentData As Double = 0
                If Me.ContainsKey(InputWordList.MemberWords(word).Sonographs_Letters(CurrentGraphemeBlockIndex)) Then
                    If Me(InputWordList.MemberWords(word).Sonographs_Letters(CurrentGraphemeBlockIndex)).ContainsKey((InputWordList.MemberWords(word).Sonographs_Pronunciation(CurrentGraphemeBlockIndex))) Then
                        CurrentData = Me(InputWordList.MemberWords(word).Sonographs_Letters(CurrentGraphemeBlockIndex))(InputWordList.MemberWords(word).Sonographs_Pronunciation(CurrentGraphemeBlockIndex)).ConditionalPredictability
                    End If
                End If

                InputWordList.MemberWords(word).G2P_OT.Add(CurrentData)

            Next

            'Calculates average spelling regularity
            CalculateAverageWordSpellingRegularity(InputWordList.MemberWords(word)) ', SetCalculationType)

        Next

    End Sub

    Public Sub CalculateAverageWordSpellingRegularity(ByRef InputWord As Word) ', ByRef SetCalculationType As GraphemeToPhonemes.CalculationTypes)

        If Not InputWord.G2P_OT.Count = 0 Then
            InputWord.G2P_OT_Average = InputWord.G2P_OT.Average
        Else
            InputWord.G2P_OT_Average = 0
        End If

    End Sub

    Public Sub Export_g2p_Data(ByVal saveDirectory As String, Optional ByRef saveFileName As String = "g2p_Data",
                                                Optional BoxTitle As String = "Choose location to store the g2p output data...",
                                       Optional ByVal SkipRounding As Boolean = False)

        'Export function for g2p data that creates a table similar to the appendix B in Berndt, R. S., Reggia, J. A., & Mitchum, C. C. (1987)

        Try

            SendInfoToLog("Sorting graphemes in alphabetic order")
            Sort_g2p_data_Alphabetically()

            SendInfoToLog("Attempting to export g2p data. SkipRounding setting: " & SkipRounding.ToString)

            'Choosing file location
            Dim filepath As String = Path.Combine(saveDirectory, saveFileName & ".txt")
            If Not Directory.Exists(Path.GetDirectoryName(filepath)) Then Directory.CreateDirectory(Path.GetDirectoryName(filepath))

            'Saving to file
            Dim writer As New StreamWriter(filepath, False, Text.Encoding.UTF8)

            'Output file structure
            'Grapheme tab Prior propbability
            'tab tab Phoneme 1 tab Conditional probability tab Example
            'tab tab Phoneme 2 tab Conditional probability tab Example
            '...
            'empty line
            'New Grapheme...

            'Writing heading
            writer.WriteLine("Grapheme" & vbTab & "Prior propbability" & vbTab & "Phoneme" & vbTab & "FrequencyData" & vbTab & "Conditional probability" & vbTab & "Conditional Predictability" & vbTab & "Example")
            writer.WriteLine("")

            writer.WriteLine("Summed frequency data: " & Me.Total_FrequencyData)
            writer.WriteLine(Generate_g2p_OutputSectionString(SkipRounding))

            writer.Close()

            SendInfoToLog("g2p data was exported to file. SkipRounding setting: " & SkipRounding.ToString & filepath)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub

    Private Function Generate_g2p_OutputSectionString(ByVal SkipRounding As Boolean) As String

        Dim OutputPreparationList As New List(Of String)

        For Each g In Me

            'Adding grapheme line to the output string
            OutputPreparationList.Add(g.Key & vbTab & Rounding(g.Value.PriorProbability,, 4, SkipRounding) & vbTab &
                                              Rounding(g.Value.FrequencyData, , 4, SkipRounding) & vbCrLf)

            'Going though each possible phoneme
            For Each p In g.Value

                Dim ExampleWordString As String = "Please add manually, no example word without unresolved g2p!"
                If Not p.Value.ExampleWord Is Nothing Then ExampleWordString = p.Value.ExampleWord.OrthographicForm &
                                                  " [" & String.Concat(p.Value.ExampleWord.BuildExtendedIpaArray(,,,,, False, False)) & "]"

                'Adding phoneme lines to the output string
                OutputPreparationList.Add(vbTab & vbTab & p.Key & vbTab &
                                                  Rounding(p.Value.FrequencyData, , 4, SkipRounding) & vbTab &
                                                  Rounding(p.Value.ConditionalProbability, , 4, SkipRounding) & vbTab &
                                                  Rounding(p.Value.ConditionalPredictability, , 4, SkipRounding) & vbTab &
                                                  ExampleWordString & vbCrLf)

            Next

            'Adding empty line to the output string
            OutputPreparationList.Add(vbCrLf)

        Next

        Dim OutputString As String = String.Concat(OutputPreparationList)

        Return OutputString

    End Function


    Public Sub Sort_g2p_data_Alphabetically()

        'Adds the graphemes to a sorted list, using the center grapheme as sort criteria
        Dim tempSortList As New SortedList(Of String, StringGraphemeCombination)

        For Each StringGraphemeComb In Me
            Dim newStringGraphemeComb As New StringGraphemeCombination
            newStringGraphemeComb.MyString = StringGraphemeComb.Key
            newStringGraphemeComb.MyGrapheme = StringGraphemeComb.Value
            Dim sortGrapheme As String = ""
            sortGrapheme = StringGraphemeComb.Key
            tempSortList.Add(sortGrapheme, newStringGraphemeComb)
        Next

        'Removes all grapheme from g2pDataToSort
        Me.Clear()

        'Adds the grapheme in alphabetically sorted order
        For Each grapheme In tempSortList
            Me.Add(grapheme.Value.MyString, grapheme.Value.MyGrapheme)
        Next

    End Sub

    ''' <summary>
    ''' Loads g2p probability data from file. NB: For performance reasons only probability data, but nothing else, is loaded.
    ''' </summary>
    ''' <param name="SetWordStartMarker"></param>
    ''' <param name="SetWordEndMarker"></param>
    ''' <param name="SetUnresolved_p2g_Character"></param>
    ''' <param name="SetCalculationUnit"></param>
    ''' <param name="FilePath"></param>
    ''' <returns></returns>
    Public Shared Function Load_g2p_DataFromFile(Optional ByRef SetWordStartMarker As String = "*", Optional ByRef SetWordEndMarker As String = "_",
                                       Optional ByVal SetUnresolved_p2g_Character As String = "!",
                           Optional ByRef SetCalculationUnit As WordGroup.WordFrequencyUnit = WordGroup.WordFrequencyUnit.WordType,
                                                       Optional ByRef FilePath As String = "") As GraphemeToPhonemes

        'NB: This function reads only the probability data from the input file! Everything else is ignored

        Dim Output As New GraphemeToPhonemes(SetWordStartMarker, SetWordEndMarker, SetUnresolved_p2g_Character, SetCalculationUnit)

        'Reading data
        Try

            Dim inputArray() As String = {}
            If FilePath = "" Then
                Dim dataString As String = My.Resources.g2p_Data
                dataString = dataString.Replace(vbCrLf, vbLf)
                inputArray = dataString.Split(vbLf)
            Else
                inputArray = System.IO.File.ReadAllLines(FilePath)
            End If

            Dim CurrentGraphemeString As String = ""

            For line = 3 To inputArray.Length - 1

                If inputArray(line).Trim(vbTab).Trim = "" Then Continue For 'This line is not quite fail safe!
                Dim LineSplit() As String = inputArray(line).Split(vbTab)

                'Check for new graphemes in the first column
                If LineSplit(0).Trim <> "" Then

                    Dim newGrapheme As New GraphemeToPhonemes.Grapheme
                    CurrentGraphemeString = LineSplit(0).Trim
                    Output.Add(CurrentGraphemeString, newGrapheme)

                Else
                    'There's nothing in the first column

                    'Checks for new phonemes in the third column
                    If LineSplit(2).Trim <> "" Then
                        Dim newPhonemeCombination As New GraphemeToPhonemes.PhonemeData
                        newPhonemeCombination.ConditionalProbability = LineSplit(4).Trim
                        newPhonemeCombination.ConditionalPredictability = LineSplit(5).Trim
                        Output(CurrentGraphemeString).Add(LineSplit(2).Trim, newPhonemeCombination)
                    End If
                End If
            Next

        Catch ex As Exception
            Console.WriteLine(ex.ToString)
            Return Nothing
        End Try

        Return Output

    End Function

End Class

''' <summary>
''' This class can be used to calculate the Grapheme-initial-letter-to-pronunciation orthographic transparency (GIL2P-OT) and the 
''' Pronunciation-initial-phone-to-grapheme orthographic transparency (PIP2G-OT) described by Witte et al. 2019. 
''' It is written in a general way so that it can be used to calculate both GIL2P-OT and PIP2G-OT, using the terminology KIS2V-OT (key-initial segment to value). 
''' </summary>
Public Class KeyInitialSegmentToValueProbability
    Inherits Dictionary(Of String, KeyInitialSegment)
    Property WordStartMarker As String
    Property WordEndMarker As String
    Property Unresolved_p2g_Character As String
    Property FrequencyData As Double
    Property OrthographicTransparencyType As OrthographicTransparencyTypes

    Public Enum OrthographicTransparencyTypes
        GIL2P
        PIP2G
    End Enum

    Public Sub New(ByVal OrthographicTransparencyType As OrthographicTransparencyTypes,
                           Optional ByRef SetWordStartMarker As String = "*", Optional ByRef SetWordEndMarker As String = "_",
                           Optional ByRef SetUnresolved_p2g_Character As String = "!")

        Me.OrthographicTransparencyType = OrthographicTransparencyType
        Unresolved_p2g_Character = SetUnresolved_p2g_Character
        WordStartMarker = SetWordStartMarker
        WordEndMarker = SetWordEndMarker

    End Sub

    Public Class KeyInitialSegment
            Inherits Dictionary(Of String, KIS2V_Key)
            Property FrequencyData As Double
            Property HighestConditionalProbability As Decimal = 0

            Public Sub SortInProbabilityOrder()

                Dim unsortedList As New List(Of KeyValuePair(Of String, KIS2V_Key))
                For Each Current_KIS2V_Key In Me
                    unsortedList.Add(New KeyValuePair(Of String, KIS2V_Key)(Current_KIS2V_Key.Key, Current_KIS2V_Key.Value))
                Next

                'Sorting in descending order
                Dim Query1 = unsortedList.OrderByDescending(Function(p) p.Value.ConditionalProbability)

                'Adding in sorted order
                Dim mySortedList As New List(Of KeyValuePair(Of String, KIS2V_Key))
                For Each p In Query1
                    mySortedList.Add(p)
                Next

                'Clearing Me
                Me.Clear()

                'Putting the data back into Me
                For Each Current_KIS2V_Key In mySortedList
                    Me.Add(Current_KIS2V_Key.Key, Current_KIS2V_Key.Value)
                Next

            End Sub
        End Class

        Public Class KIS2V_Key
            Inherits Dictionary(Of String, KIS2V_Value)

            Property FrequencyData As Double
            Property ConditionalProbability As Decimal = 0
            Property HighestConditionalProbability As Decimal = 0

            Public Sub SortInProbabilityOrder()

                Dim unsortedList As New List(Of KeyValuePair(Of String, KIS2V_Value))
                For Each Current_KIS2V_Value In Me
                    unsortedList.Add(New KeyValuePair(Of String, KIS2V_Value)(Current_KIS2V_Value.Key, Current_KIS2V_Value.Value))
                Next

                'Sorting in descending order
                Dim Query1 = unsortedList.OrderByDescending(Function(p) p.Value.g2p_Predictability)

                'Adding in sorted order
                Dim mySortedList As New List(Of KeyValuePair(Of String, KIS2V_Value))
                For Each p In Query1
                    mySortedList.Add(p)
                Next

                'Clearing Me
                Me.Clear()

                'Putting the data back into Me
                For Each Current_KIS2V_Key In mySortedList
                    Me.Add(Current_KIS2V_Key.Key, Current_KIS2V_Key.Value)
                Next
            End Sub
        End Class


        Public Class KIS2V_Value

            Property FrequencyData As Double
            Property Conditional_gb2p_Probability As Decimal = 0
            Property Conditional_g2p_Probability As Decimal = 0
            Property g2p_Predictability As Decimal = 0
            Property ExampleWord As Word

        End Class





        Private Sub AddData(ByRef CurrentWord As Word, ByRef KIS As String, ByRef KIS_Key As String, ByRef KIS_Value As String,
                                       Optional ByRef ExampleWordFrequencyLimit As Integer = 10,
                                      Optional ByRef DoNotAddDataWithZeroFrequency As Boolean = True,
                                      Optional ByRef DoNotAddDataWithUnresolved_p2g As Boolean = True)

            Dim FrequencyData As Double = 0
            FrequencyData = CurrentWord.RawWordTypeFrequency

            'Not adding if DoNotAddDataWithZeroFrequency = true and the Frequency value is zero
            If DoNotAddDataWithZeroFrequency = True And FrequencyData = 0 Then Exit Sub

            'Not adding if it is an unresolved p2g
            If DoNotAddDataWithUnresolved_p2g = True Then
                If KIS_Key.Contains(Unresolved_p2g_Character) Or KIS_Value.Contains(Unresolved_p2g_Character) Then
                    Exit Sub
                End If
            End If

            'Checking if the KIS exists
            If Not Me.ContainsKey(KIS) Then
                'Adds level 1
                Dim newKeyInitialSegment As New KeyInitialSegment
                Me.Add(KIS, newKeyInitialSegment)
            End If

            'Adds level 2
            If Not Me(KIS).ContainsKey(KIS_Key) Then
                Dim newKIS2V_Key As New KIS2V_Key
                Me(KIS).Add(KIS_Key, newKIS2V_Key)
            End If

            'Adds level 3
            If Not Me(KIS)(KIS_Key).ContainsKey(KIS_Value) Then
                Dim newKIS2V_Value As New KIS2V_Value
                Me(KIS)(KIS_Key).Add(KIS_Value, newKIS2V_Value)
                Me(KIS)(KIS_Key)(KIS_Value).FrequencyData = FrequencyData

                'Example words
                'Not adding Hapax legomenas
                If Not CurrentWord.RawWordTypeFrequency >= ExampleWordFrequencyLimit Then
                    'Not adding example words if they contain an unresolved character (as this may acctually cause erroneous segmentation)
                    If Not String.Concat(CurrentWord.Sonographs_Letters).Contains(Unresolved_p2g_Character) Then
                        Me(KIS)(KIS_Key)(KIS_Value).ExampleWord = CurrentWord
                    End If
                End If

            Else

                'Increases frequency value
                Me(KIS)(KIS_Key)(KIS_Value).FrequencyData += FrequencyData

                'Example word selection
                'Not adding Hapax legomenas
                If Not CurrentWord.RawWordTypeFrequency >= ExampleWordFrequencyLimit Then

                    'Not adding example words if they contain an unresolved character (as this may acctually cause erroneous segmentation)
                    If Not String.Concat(CurrentWord.Sonographs_Letters).Contains(Unresolved_p2g_Character) Then

                        'Adding the word if no word has been previously added
                        If Me(KIS)(KIS_Key)(KIS_Value).ExampleWord Is Nothing Then
                            Me(KIS)(KIS_Key)(KIS_Value).ExampleWord = CurrentWord
                        Else

                            'Replacing the existing word if it is marked as foreign, and the new one is not
                            If Me(KIS)(KIS_Key)(KIS_Value).ExampleWord.ForeignWord = True And CurrentWord.ForeignWord = False Then
                                Me(KIS)(KIS_Key)(KIS_Value).ExampleWord = CurrentWord
                            Else

                                'Replacing the existing word if contains a special character, and new one is not
                                If Me(KIS)(KIS_Key)(KIS_Value).ExampleWord.OrthographicFormContainsSpecialCharacter = True And CurrentWord.OrthographicFormContainsSpecialCharacter = False Then
                                    Me(KIS)(KIS_Key)(KIS_Value).ExampleWord = CurrentWord
                                Else

                                    'Replaces the existing word if it has a phoneme range of 3-8 phonemes and a higher phonotactic probability
                                    Dim PhonemeCount As Integer = CurrentWord.CountPhonemes
                                    If PhonemeCount > 2 And PhonemeCount < 9 Then

                                        Dim NewWordValues As New List(Of Double) From {CurrentWord.SSPP_Average, CurrentWord.RawWordTypeFrequency}
                                        Dim ExistingWordValues As New List(Of Double) From {Me(KIS)(KIS_Key)(KIS_Value).ExampleWord.SSPP_Average,
                                            Me(KIS)(KIS_Key)(KIS_Value).ExampleWord.RawWordTypeFrequency}

                                        If GeometricMean(NewWordValues) > GeometricMean(ExistingWordValues) Then
                                            Me(KIS)(KIS_Key)(KIS_Value).ExampleWord = CurrentWord
                                        End If

                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If


        End Sub

        ''' <summary>
        ''' Returns the initial letter to phoneme predictability, or 0 if look-up fails due to missing initial grapheme letter, grapheme key, or phoneme in the probability matrix.
        ''' </summary>
        ''' <returns></returns>
        Private Function Get_Data(ByRef KIS As String, ByRef KIS_Key As String, ByRef KIS_Value As String)

            Dim Output As Double = 0

            'Returns 0 for transition to silent graphemes in PIP2G-OT
            If OrthographicTransparencyType = OrthographicTransparencyTypes.PIP2G Then
                If KIS = ZeroPhoneme Then Return 0
            End If

            If Me.ContainsKey(KIS) Then
                If Me(KIS).ContainsKey(KIS_Key) Then
                    If Me(KIS)(KIS_Key).ContainsKey(KIS_Value) Then
                        Output = Me(KIS)(KIS_Key)(KIS_Value).g2p_Predictability
                    End If
                End If
            End If

            'In PIP2g OT, zero-phones should really get a probability of 0, since they can hardly be said to be predictable.

            Return Output

        End Function


        ''' <summary>
        ''' Using the collected KIS2V data, probability data for each KIS to V correspondence is calculated
        ''' </summary>
        Private Sub CalculateProbabilityData()

            SendInfoToLog("Initializes calculation of KIS2V probability.")
            Dim ZeroProbabilityCount As Integer = 0
            Dim ZeroPredictabilityCount As Integer = 0

            'Margin to be used in testing probability sum
            Dim RoundingMargin As Double = 0.00001

            'Sums frequency data, on the different levels
            Me.FrequencyData = 0
            For Each Current_KIS As KeyValuePair(Of String, KeyInitialSegment) In Me
                Current_KIS.Value.FrequencyData = 0
                For Each Current_KIS2V_Key In Current_KIS.Value
                    Current_KIS2V_Key.Value.FrequencyData = 0
                    For Each Current_KIS2V_Value In Current_KIS2V_Key.Value

                        'Modifies some data slightly so that no extant transitions gets a zero probability./ 
                        Current_KIS2V_Value.Value.FrequencyData = Math.Log10(Current_KIS2V_Value.Value.FrequencyData + 1) 'The term +1 is used to avoid that FrequencyData of 1 becomes a zero probability: log10(1) = 0. This means that the phoneme count is very slightly elevated, by 1

                        'Sums log values
                        Current_KIS2V_Key.Value.FrequencyData += Current_KIS2V_Value.Value.FrequencyData

                    Next

                    'Sums non-log values
                    Current_KIS.Value.FrequencyData += Current_KIS2V_Key.Value.FrequencyData
                Next
                Me.FrequencyData += Current_KIS.Value.FrequencyData
            Next


            'B. Calculating conditional probabilities of each "grapheme-block start grapheme" to its member grapheme blocks
            For Each Current_KIS As KeyValuePair(Of String, KeyInitialSegment) In Me
                Dim ProbCheckSum As Double = 0

                For Each Current_KIS2V_Key In Current_KIS.Value
                    'Calculating conditional grapheme-to-grapheme block probabilities

                    Current_KIS2V_Key.Value.ConditionalProbability = Current_KIS2V_Key.Value.FrequencyData / Current_KIS.Value.FrequencyData

                    'Accumulating test probability data
                    ProbCheckSum += Current_KIS2V_Key.Value.ConditionalProbability
                Next

                'Checking if probabilties add up to 1
                If ProbCheckSum < (1 - RoundingMargin) Or ProbCheckSum > (1 + RoundingMargin) Then
                    MsgBox("Probabilties do not add up to 1")
                End If

            Next


            'C. Calculating conditional probabilities of each grapheme block to its member phonemes
            For Each Current_KIS As KeyValuePair(Of String, KeyInitialSegment) In Me
                For Each Current_KIS2V_Key In Current_KIS.Value

                    Dim ProbCheckSum As Double = 0
                    For Each Current_KIS2V_Value In Current_KIS2V_Key.Value

                        'Calculating conditional grapheme block-to-phoneme probabilities
                        Current_KIS2V_Value.Value.Conditional_gb2p_Probability = Current_KIS2V_Value.Value.FrequencyData / Current_KIS2V_Key.Value.FrequencyData

                        'Accumulating test probability data
                        ProbCheckSum += Current_KIS2V_Value.Value.Conditional_gb2p_Probability

                    Next

                    'Checking if probabilties add up to 1
                    If ProbCheckSum < (1 - RoundingMargin) Or ProbCheckSum > (1 + RoundingMargin) Then
                        MsgBox("Probabilties do not add up to 1")
                    End If

                Next
            Next

            'D. Calculating g2p probability by multiplying the probabilities under B with those under C
            For Each GraphemeStart As KeyValuePair(Of String, KeyInitialSegment) In Me

                Dim ProbCheckSum As Double = 0

                For Each Current_KIS2V_Key In GraphemeStart.Value
                    For Each Current_KIS2V_Value In Current_KIS2V_Key.Value

                        'Calculating g2p probability 
                        Current_KIS2V_Value.Value.Conditional_g2p_Probability = Current_KIS2V_Key.Value.ConditionalProbability * Current_KIS2V_Value.Value.Conditional_gb2p_Probability

                        'Accumulating test probability data
                        ProbCheckSum += Current_KIS2V_Value.Value.Conditional_g2p_Probability

                    Next
                Next

                'Checking if probabilties add up to 1
                If ProbCheckSum < (1 - RoundingMargin) Or ProbCheckSum > (1 + RoundingMargin) Then
                    MsgBox("Probabilties do not add up to 1")
                End If

            Next

            'I. Clearing the probability matrix from 0 valued probabilities
            SendInfoToLog("Removing zero probability l2p transitions.")
            Dim KIS2V_ValuesRemoved As Integer = 0
            For Each Current_KIS As KeyValuePair(Of String, KeyInitialSegment) In Me
                For Each Current_KIS2V_Key In Current_KIS.Value

                    Dim Current_KIS2V_ValuesToRemove As New List(Of String)
                    For Each Current_KIS2V_Value In Current_KIS2V_Key.Value
                        If Current_KIS2V_Value.Value.Conditional_g2p_Probability = 0 Then
                            If Not Current_KIS2V_ValuesToRemove.Contains(Current_KIS2V_Value.Key) Then Current_KIS2V_ValuesToRemove.Add(Current_KIS2V_Value.Key)
                        End If
                    Next

                    For Each Current_KIS2V_Value In Current_KIS2V_ValuesToRemove
                        Me(Current_KIS.Key)(Current_KIS2V_Key.Key).Remove(Current_KIS2V_Value)
                        KIS2V_ValuesRemoved += 1
                    Next
                Next
            Next

            'Clearing empty graphemeblocks (that became empty in the preceding step)
            Dim KIS2V_KeysRemoved As Integer = 0
            For Each Current_KIS As KeyValuePair(Of String, KeyInitialSegment) In Me

                Dim Current_KIS2V_KeysToRemove As New List(Of String)
                For Each Current_KIS2V_Key In Current_KIS.Value
                    If Current_KIS2V_Key.Value.Count = 0 Then
                        If Not Current_KIS2V_KeysToRemove.Contains(Current_KIS2V_Key.Key) Then Current_KIS2V_KeysToRemove.Add(Current_KIS2V_Key.Key)
                    End If
                Next

                For Each Current_KIS2V_Key In Current_KIS2V_KeysToRemove
                    Me(Current_KIS.Key).Remove(Current_KIS2V_Key)
                    KIS2V_KeysRemoved += 1
                Next
            Next

            'Clearing empty grapheme starts
            Dim KISsRemoved As Integer = 0
            Dim KISsToRemove As New List(Of String)
            For Each Current_KIS As KeyValuePair(Of String, KeyInitialSegment) In Me
                If Current_KIS.Value.Count = 0 Then
                    If Not KISsToRemove.Contains(Current_KIS.Key) Then KISsToRemove.Add(Current_KIS.Key)
                End If
            Next
            For Each CurrentGraphemeStart In KISsToRemove
                Me.Remove(CurrentGraphemeStart)
                KISsRemoved += 1
            Next

            SendInfoToLog("Finished removing zero probability l2p transitions. Results: " & KIS2V_ValuesRemoved & " KIS2V_Values and " & KIS2V_KeysRemoved & " KIS2V_Keys and " & KISsRemoved & " KISs were removed.")


            'Determining the highest g2p conditional probability for each grapheme start
            For Each Current_KIS As KeyValuePair(Of String, KeyInitialSegment) In Me
                Current_KIS.Value.HighestConditionalProbability = 0
                For Each Current_KIS2V_Key In Current_KIS.Value
                    For Each Current_KIS2V_Value In Current_KIS2V_Key.Value

                        'Detecting the highest probability for the current GraphemeStart
                        If Current_KIS2V_Value.Value.Conditional_g2p_Probability > Current_KIS.Value.HighestConditionalProbability Then
                            Current_KIS.Value.HighestConditionalProbability = Current_KIS2V_Value.Value.Conditional_g2p_Probability
                        End If

                    Next
                Next
            Next

            'Calculating gil2p predictability / normalized gil2p probability
            For Each Current_KIS As KeyValuePair(Of String, KeyInitialSegment) In Me
                For Each Current_KIS2V_Key In Current_KIS.Value
                    For Each Current_KIS2V_Value In Current_KIS2V_Key.Value
                        Current_KIS2V_Value.Value.g2p_Predictability = Current_KIS2V_Value.Value.Conditional_g2p_Probability / Current_KIS.Value.HighestConditionalProbability
                    Next
                Next
            Next


            'H. Setting the Predictabilities and probabilties for unresolved p2gs to 0
            If Me.ContainsKey(Unresolved_p2g_Character) Then
                'Me(Unresolved_p2g_Character).TW_FrequencyData = 0
                Me(Unresolved_p2g_Character).HighestConditionalProbability = 0

                For Each Current_KIS2V_Key In Me(Unresolved_p2g_Character)
                    Current_KIS2V_Key.Value.ConditionalProbability = 0
                    Current_KIS2V_Key.Value.HighestConditionalProbability = 0

                    For Each Current_KIS2V_Value In Current_KIS2V_Key.Value
                        Current_KIS2V_Value.Value.Conditional_g2p_Probability = 0
                        Current_KIS2V_Value.Value.Conditional_gb2p_Probability = 0
                        Current_KIS2V_Value.Value.g2p_Predictability = 0
                    Next
                Next
            End If

            SendInfoToLog("Finished calculating KIS2P probabilities. Results: " & vbCrLf &
                              " ZeroPredictabilityCount encounterred: " & ZeroProbabilityCount & vbCrLf &
                              " ZeroPredictabilityCount encounterred: " & ZeroPredictabilityCount)

        End Sub


        Public Sub Collect_KIS2V_OT_Data(ByRef InputWordGroup As WordGroup, ByRef OutputFolder As String)

            Try

                'Clears all data
                Me.Clear()

                SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

                'Going through all member words and collect frequency data, and stores it in fb2gData
                For word = 0 To InputWordGroup.MemberWords.Count - 1

                    Dim TempSpelling As String = Me.WordStartMarker & InputWordGroup.MemberWords(word).OrthographicForm & Me.WordEndMarker

                    For CurrentGraphemeBlockIndex = 0 To InputWordGroup.MemberWords(word).Sonographs_Letters.Count - 1

                        'Adding OT Data
                        Try

                            Dim CurrentKIS As String = ""
                            Dim CurrentKIS2V_Key As String = ""
                            Dim CurrentKIS2V_Value As String = ""

                            Select Case OrthographicTransparencyType
                                Case OrthographicTransparencyTypes.GIL2P

                                    CurrentKIS2V_Key = InputWordGroup.MemberWords(word).Sonographs_Letters(CurrentGraphemeBlockIndex)
                                    CurrentKIS = CurrentKIS2V_Key.Substring(0, 1)
                                    CurrentKIS2V_Value = InputWordGroup.MemberWords(word).Sonographs_Pronunciation(CurrentGraphemeBlockIndex)

                                Case OrthographicTransparencyTypes.PIP2G

                                    CurrentKIS2V_Key = InputWordGroup.MemberWords(word).Sonographs_Pronunciation(CurrentGraphemeBlockIndex)
                                    CurrentKIS = CurrentKIS2V_Key.Split(" ")(0)
                                    CurrentKIS2V_Value = InputWordGroup.MemberWords(word).Sonographs_Letters(CurrentGraphemeBlockIndex)

                            End Select

                            Me.AddData(InputWordGroup.MemberWords(word), CurrentKIS, CurrentKIS2V_Key, CurrentKIS2V_Value)

                        Catch ex As Exception
                            MsgBox(ex.ToString)
                        End Try

                    Next
                Next

                'Calculates spelling probability
                Me.CalculateProbabilityData()

                SendInfoToLog("Method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " completed successfully. Data was collected from " & InputWordGroup.MemberWords.Count & " words.")

                'Exporting data
                Select Case OrthographicTransparencyType
                    Case OrthographicTransparencyTypes.GIL2P
                        Export_DataFile(OutputFolder, "GIL2P_Data",,,, True)
                        Export_DataFile(OutputFolder, "GIL2P_Data_Rounded",,,, False)

                    Case OrthographicTransparencyTypes.PIP2G
                        Export_DataFile(OutputFolder, "PIP2G_Data",,,, True)
                        Export_DataFile(OutputFolder, "PIP2G_Data_Rounded",,,, False)

                End Select

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End Sub


        Public Sub Apply_KIS2V_OT_Data(ByRef InputWordGroup As WordGroup)

            Try

                'Going through all member words and looks up their spelling probability values
                For word = 0 To InputWordGroup.MemberWords.Count - 1

                    'Resetting spellingRegularity
                    Select Case OrthographicTransparencyType
                        Case OrthographicTransparencyTypes.GIL2P
                            InputWordGroup.MemberWords(word).GIL2P_OT.Clear()
                        Case OrthographicTransparencyTypes.PIP2G
                            InputWordGroup.MemberWords(word).PIP2G_OT.Clear()
                    End Select

                    Dim TempSpelling As String = Me.WordStartMarker & InputWordGroup.MemberWords(word).OrthographicForm & WordEndMarker
                    Dim OrthStartIndex As Integer = 0

                    For CurrentGraphemeBlockIndex = 0 To InputWordGroup.MemberWords(word).Sonographs_Letters.Count - 1

                        'Applying OT data to the current word
                        Try

                            Dim CurrentKIS As String = ""
                            Dim CurrentKIS2V_Key As String = ""
                            Dim CurrentKIS2V_Value As String = ""

                            Select Case OrthographicTransparencyType
                                Case OrthographicTransparencyTypes.GIL2P

                                    CurrentKIS2V_Key = InputWordGroup.MemberWords(word).Sonographs_Letters(CurrentGraphemeBlockIndex)
                                    CurrentKIS = CurrentKIS2V_Key.Substring(0, 1)
                                    CurrentKIS2V_Value = InputWordGroup.MemberWords(word).Sonographs_Pronunciation(CurrentGraphemeBlockIndex)

                                Case OrthographicTransparencyTypes.PIP2G

                                    CurrentKIS2V_Key = InputWordGroup.MemberWords(word).Sonographs_Pronunciation(CurrentGraphemeBlockIndex)
                                    CurrentKIS = CurrentKIS2V_Key.Split(" ")(0)
                                    CurrentKIS2V_Value = InputWordGroup.MemberWords(word).Sonographs_Letters(CurrentGraphemeBlockIndex)

                            End Select

                            Select Case OrthographicTransparencyType
                                Case OrthographicTransparencyTypes.GIL2P

                                    InputWordGroup.MemberWords(word).GIL2P_OT.Add(Get_Data(CurrentKIS, CurrentKIS2V_Key, CurrentKIS2V_Value))

                                Case OrthographicTransparencyTypes.PIP2G

                                    InputWordGroup.MemberWords(word).PIP2G_OT.Add(Get_Data(CurrentKIS, CurrentKIS2V_Key, CurrentKIS2V_Value))

                            End Select


                        Catch ex As Exception
                            MsgBox(ex.ToString)
                        End Try

                    Next

                    'Calculates average orthographic transparency
                    CalculateAverageWordOrthographicTransparency(InputWordGroup.MemberWords(word))
                    CalculateMinimumWordOrthographicTransparency(InputWordGroup.MemberWords(word))

                Next


                SendInfoToLog("Method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " completed successfully. Results: Spelling regularity was calculated for " & InputWordGroup.MemberWords.Count & " words.")

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End Sub


        Private Sub CalculateAverageWordOrthographicTransparency(ByRef InputWord As Word)

            Select Case OrthographicTransparencyType
                Case OrthographicTransparencyTypes.GIL2P

                    If Not InputWord.GIL2P_OT.Count = 0 Then
                        InputWord.GIL2P_OT_Average = InputWord.GIL2P_OT.Average
                    Else
                        InputWord.GIL2P_OT_Average = 0
                    End If

                Case OrthographicTransparencyTypes.PIP2G

                    If Not InputWord.PIP2G_OT.Count = 0 Then
                        InputWord.PIP2G_OT_Average = InputWord.PIP2G_OT.Average
                    Else
                        InputWord.PIP2G_OT_Average = 0
                    End If

            End Select


        End Sub

        Private Sub CalculateMinimumWordOrthographicTransparency(ByRef InputWord As Word)

            Select Case OrthographicTransparencyType
                Case OrthographicTransparencyTypes.GIL2P

                    If Not InputWord.GIL2P_OT.Count = 0 Then
                        InputWord.GIL2P_OT_Min = InputWord.GIL2P_OT.Min
                    Else
                        InputWord.GIL2P_OT_Min = 0
                    End If

                Case OrthographicTransparencyTypes.PIP2G

                    If Not InputWord.PIP2G_OT.Count = 0 Then
                        InputWord.PIP2G_OT_Min = InputWord.PIP2G_OT.Min
                    Else
                        InputWord.PIP2G_OT_Min = 0
                    End If

            End Select


        End Sub

        Private Sub Sort_KIS2V_data()

            'Sorts level 1 alphabetically
            Dim tempSortList_Level1 As New SortedList(Of String, KeyInitialSegment)
            For Each CurrentKIS In Me
                tempSortList_Level1.Add(CurrentKIS.Key, CurrentKIS.Value)
            Next
            Me.Clear()
            For Each CurrentGrapheme In tempSortList_Level1
                Me.Add(CurrentGrapheme.Key, CurrentGrapheme.Value)
            Next

            'Sorts level 2 in probability order
            For Each Current_KIS2V_Key In Me
                Current_KIS2V_Key.Value.SortInProbabilityOrder()
            Next

            'Sorts level 3 in probability order
            For Each CurrentGrapheme In Me
                For Each CurrentGraphemeBlock In CurrentGrapheme.Value
                    CurrentGraphemeBlock.Value.SortInProbabilityOrder()
                Next
            Next

        End Sub



        Public Sub Export_DataFile(Byval saveDirectory As String , Optional ByRef saveFileName As String = "KIS2V_Data",
                                                Optional BoxTitle As String = "Choose location to store the KIS2V output data...",
                                           Optional ByRef ExcludeEmptyFields As Boolean = True,
                                           Optional ByRef SortResults As Boolean = True,
                                         Optional ByVal SkipRounding As Boolean = False)

            Try

                SendInfoToLog("Sorting KIS2V data")
                If SortResults = True Then
                    Sort_KIS2V_data()
                End If

                SendInfoToLog("Attempting to export KIS2V data. SkipRounding setting: " & SkipRounding.ToString)

            'Choosing file location
            Dim filepath As String = Path.Combine(saveDirectory, saveFileName & ".txt")
            If Not Directory.Exists(Path.GetDirectoryName(filepath)) Then Directory.CreateDirectory(Path.GetDirectoryName(filepath))


            'Saving to file
            Dim writer As New StreamWriter(filepath, False, Text.Encoding.UTF8)

                'Writing heading
                writer.WriteLine("Grapheme" & vbTab & "FreqData" & vbTab & "HighestProb" & vbTab &
                                         "Grapheme block" & vbTab & "FreqData" & vbTab & "KIS2K_Conditional propbability" & vbTab &
                                         "Phoneme" & vbTab & "FreqData" & vbTab & "K2V_Conditional probability" & vbTab &
                                         "KIS2V_Conditional probability" & vbTab & "KIS2V_Predictability" & vbTab & "Examples")


                writer.WriteLine("")

                writer.WriteLine("Total KIS count: " & Me.Count)
                writer.WriteLine(Generate_KIS2V_OutputSectionString(ExcludeEmptyFields, SkipRounding))

                writer.Close()

                SendInfoToLog("KIS2V data was exported to file: " & filepath)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try


        End Sub

    Public Shared Function Load_KIS2V_FromTxtFile(ByVal OrthographicTransparencyType As OrthographicTransparencyTypes,
                                                          Optional ByRef filePath As String = "",
                                                          Optional ByRef SetWordStartMarker As String = "*",
                                                          Optional ByRef SetWordEndMarker As String = "_",
                                                          Optional ByRef SetUnresolved_p2g_Character As String = "!") As KeyInitialSegmentToValueProbability
        'Creating an output
        Dim output As New KeyInitialSegmentToValueProbability(OrthographicTransparencyType, SetWordStartMarker, SetWordEndMarker, SetUnresolved_p2g_Character)

        'Reading data
        Try

            Dim inputArray() As String = {}

            If filePath = "" Then

                'Getting the appropriate data from My.Resources
                Dim dataString As String = ""
                Select Case OrthographicTransparencyType
                    Case OrthographicTransparencyTypes.GIL2P
                        dataString = My.Resources.GIL2P_Data
                    Case OrthographicTransparencyTypes.PIP2G
                        dataString = My.Resources.PIP2G_Data
                End Select

                dataString = dataString.Replace(vbCrLf, vbLf)
                inputArray = dataString.Split(vbLf)

            Else
                inputArray = System.IO.File.ReadAllLines(filePath)
            End If


            Dim CurrentKISString As String = ""
            Dim CurrentKIS2V_Key_String As String = ""

            For line = 3 To inputArray.Length - 1

                If inputArray(line).Trim(vbTab).Trim = "" Then Continue For 'This line is not quite fail safe!

                Dim LineSplit() As String = inputArray(line).Split(vbTab)

                'Check for new graphemes in the first column
                If LineSplit(0).Trim <> "" Then

                    Dim newKIS As New KeyInitialSegment
                    newKIS.FrequencyData = LineSplit(2).Trim
                    CurrentKISString = LineSplit(0).Trim
                    output.Add(CurrentKISString, newKIS)

                    'TODO: This should also read TW_SimpleCount etc

                Else
                    'There's nothing in the first column

                    'Check for new graphemeblocks in the fourth column
                    If LineSplit(3).Trim <> "" Then

                        Dim newKIS2V_Key As New KIS2V_Key
                        newKIS2V_Key.ConditionalProbability = LineSplit(4).Trim
                        newKIS2V_Key.FrequencyData = LineSplit(5).Trim
                        CurrentKIS2V_Key_String = LineSplit(3).Trim
                        output(CurrentKISString).Add(CurrentKIS2V_Key_String, newKIS2V_Key)

                        'TODO: This should also read TW_SimpleCount etc

                    Else
                        'There's nothing in the fourth column

                        'Check for new phonemes in the seventh column
                        If LineSplit(6).Trim <> "" Then

                            Dim newKIS2V_Value As New KIS2V_Value
                            newKIS2V_Value.FrequencyData = LineSplit(7).Trim
                            newKIS2V_Value.Conditional_gb2p_Probability = LineSplit(8).Trim
                            newKIS2V_Value.Conditional_g2p_Probability = LineSplit(9).Trim
                            newKIS2V_Value.g2p_Predictability = LineSplit(10).Trim

                            output(CurrentKISString)(CurrentKIS2V_Key_String).Add(LineSplit(6).Trim, newKIS2V_Value)

                            'TODO: This should also read TW_SimpleCount etc

                        End If
                    End If
                End If
            Next


        Catch ex As Exception
            Return Nothing
        End Try

        Return output

    End Function



    Private Function Generate_KIS2V_OutputSectionString(ByRef ExcludeNonExtantFields As Boolean, ByVal SkipRounding As Boolean) As String

            Dim OutputPreparationList As New List(Of String)

            For Each KIS In Me

                'Skipping graphemes that have no occurrences
                If KIS.Value.FrequencyData > 0 Or ExcludeNonExtantFields = False Then

                    'Adding grapheme line to the output string
                    OutputPreparationList.Add(KIS.Key & vbTab & Rounding(KIS.Value.FrequencyData, , 4, SkipRounding) & vbTab &
                                                  Rounding(KIS.Value.HighestConditionalProbability,, 4, SkipRounding) & vbCrLf)

                    'Going though each possible grapheme block
                    For Each Current_KIS2V_Key In KIS.Value

                        'Skipping grapheme blocks that have no occurrences
                        If Current_KIS2V_Key.Value.FrequencyData > 0 Or ExcludeNonExtantFields = False Then

                            'Adding grapheme block lines to the output string
                            OutputPreparationList.Add(vbTab & vbTab & vbTab &
                                                          Current_KIS2V_Key.Key & vbTab & Rounding(Current_KIS2V_Key.Value.FrequencyData, , 4, SkipRounding) & vbTab &
                                                          Rounding(Current_KIS2V_Key.Value.ConditionalProbability,, 4, SkipRounding) & vbCrLf)

                            'Going though each possible phoneme
                            For Each Current_KIS2V_Value In Current_KIS2V_Key.Value

                                'Skipping phonemes that have no occurrences
                                If Current_KIS2V_Value.Value.FrequencyData > 0 Or ExcludeNonExtantFields = False Then

                                    'Adding phoneme lines to the output string
                                    Dim ExampleWordString As String = "Please add manually, no example word without unresolved g2p!"
                                    If Not Current_KIS2V_Value.Value.ExampleWord Is Nothing Then ExampleWordString = Current_KIS2V_Value.Value.ExampleWord.OrthographicForm & " [" &
                                                                  String.Concat(Current_KIS2V_Value.Value.ExampleWord.BuildExtendedIpaArray(,,,,, False, False)) & "]"

                                    OutputPreparationList.Add(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab &
                                                                  Current_KIS2V_Value.Key & vbTab & Rounding(Current_KIS2V_Value.Value.FrequencyData,, 4, SkipRounding) & vbTab &
                                                                  Rounding(Current_KIS2V_Value.Value.Conditional_gb2p_Probability,, 4, SkipRounding) & vbTab &
                                                                  Rounding(Current_KIS2V_Value.Value.Conditional_g2p_Probability,, 4, SkipRounding) & vbTab &
                                                                  Rounding(Current_KIS2V_Value.Value.g2p_Predictability,, 4, SkipRounding) & vbTab &
                                                                  ExampleWordString & vbCrLf)

                                End If
                            Next
                        End If
                    Next

                    'Adding empty line to the output string
                    OutputPreparationList.Add(vbCrLf)

                End If

            Next

            Dim OutputString As String = String.Concat(OutputPreparationList)

            Return OutputString

        End Function


    End Class



