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


Public Class OrthographicIsolationPoints

    Public ReadOnly Property ComparisonCorpus As WordGroup.OrthographicComparisonCorpus

    Public Sub New(ByVal SpeechDataLocation As String)

        'Loading a comparison corpus from file
        ComparisonCorpus = WordGroup.OrthographicComparisonCorpus.LoadComparisonCorpus(System.IO.Path.Combine(SpeechDataLocation, "OLDComparisonCorpus_AfcList.txt"))

    End Sub

    Public Sub New(ByRef ComparisonCorpus As WordGroup.OrthographicComparisonCorpus)

        'Referencing the comparison corpus
        Me.ComparisonCorpus = ComparisonCorpus

    End Sub


    ''' <summary>
    ''' Gets the (zero-based) index of the letter by which the word can be uniquely discriminated from all other words in the comparison corpus
    ''' </summary>
    ''' <returns></returns>
    Public Function GetIsolationPoint(ByVal InputSpelling As String, Optional ByRef ErrorString As String = "") As Integer

        Try

            'Go through each letter index and exclude words with different letters at that index, stop when no contrasting words exist


            'Checks if any contrasting words are added, if not 0 is returned since discrimination will be possible at the first letter
            If ComparisonCorpus.Count = 0 Then Return 0

            Dim CurrentComparisonCorpus As New List(Of String)

            'Stores all words for which the initial letter agrees with the InputSpelling in a new list
            For Each ComparisonWord In ComparisonCorpus

                'Checks if its the same word, the just skips to next
                If InputSpelling = ComparisonWord.Key Then Continue For

                'Checks if the current letter is different
                If InputSpelling.Substring(0, 1) = ComparisonWord.Key.Substring(0, 1) Then

                    'Adds the contrasting word (which should have exactly the same word initial letter sequence)
                    CurrentComparisonCorpus.Add(ComparisonWord.Key)
                End If
            Next

            'Checks if any contrasting words are added, if not 0 is returned 
            If CurrentComparisonCorpus.Count = 0 Then Return 0


            'Goes through the remaining positions in the spelling and stores all words with identical sequences so far in CurrentWordLengthComparisonCorpus
            For i = 1 To InputSpelling.Length - 1

                Dim CurrentWordLengthComparisonCorpus As New List(Of String)

                'Ignores any contrasting spellings that have different or no letters at the current sub-string index
                For Each ComparisonSpelling In CurrentComparisonCorpus

                    'Checks if its the same word, the just skips to next
                    If InputSpelling = ComparisonSpelling Then Continue For

                    'Checks if the word end is reached
                    If i >= ComparisonSpelling.Length Then Continue For

                    'Checks if the current letter is different
                    If InputSpelling.Substring(i, 1) <> ComparisonSpelling.Substring(i, 1) Then Continue For

                    'Adds the contrasting word (which should have exactly the same word initial letter sequence)
                    CurrentWordLengthComparisonCorpus.Add(ComparisonSpelling)

                Next

                'Replacing the CurrentComparisonCoprus by the CurrentWordLengthComparisonCorpus in order to only compare identical items (so far) in the next step
                CurrentComparisonCorpus = CurrentWordLengthComparisonCorpus

                'Checks if any contrasting words are added, if not i is returned 
                If CurrentComparisonCorpus.Count = 0 Then Return i

            Next

            'Returns the position after the last index, as other words with an initial identical letter sequence but are longer exist
            Return InputSpelling.Length

        Catch ex As Exception
            ErrorString &= ex.ToString & vbCrLf
            Return -1
        End Try

    End Function


End Class





Public Class PhoneticIsolationPoints

    Public ReadOnly Property ComparisonCorpus As WordGroup.PhoneticComparisonCorpus

    Public Sub New(ByVal SpeechDataLocation As String)

        'Loading a comparison corpus from file
        Dim PLD1Corpus = WordGroup.PLDComparisonCorpus.LoadComparisonCorpus(System.IO.Path.Combine(SpeechDataLocation, "PLDComparisonCorpus_AfcList.txt"))
        ComparisonCorpus = ConvertPLD1CorpusToComparisonCorpus(PLD1Corpus)

    End Sub

    Public Sub New(ByRef ComparisonCorpus As WordGroup.PhoneticComparisonCorpus)

        'Referencing the comparison corpus
        Me.ComparisonCorpus = ComparisonCorpus

    End Sub

    Public Sub New(ByRef ComparisonCorpus As WordGroup.PLDComparisonCorpus)

        'Referencing the comparison corpus
        Me.ComparisonCorpus = ConvertPLD1CorpusToComparisonCorpus(ComparisonCorpus)

    End Sub


    Private Function ConvertPLD1CorpusToComparisonCorpus(ByRef PLD1ComparisonCorpus As WordGroup.PLDComparisonCorpus) As WordGroup.PhoneticComparisonCorpus

        'Removing the syllable length info and collapsing items into one list. 
        'Primary stress syllable index and tone is also removed (since these are supra-segmental properties). 
        'More than one word may with identical transcriptions may be identified, for example in the case of minimap pairs for tone.

        Dim Output As New WordGroup.PhoneticComparisonCorpus

        For Each SyllableCountData In PLD1ComparisonCorpus
            For Each PLD1_Transcription In SyllableCountData.Value

                Dim Transcription As New List(Of String)
                'Adding the items starting on index 2. Index 0 is the pitch accent/tone, and index 1 is the main stress syllable index

                For PhoneIndex = 2 To PLD1_Transcription.Value.PLD1Transcription.Count - 1
                    Transcription.Add(PLD1_Transcription.Value.PLD1Transcription(PhoneIndex))
                Next
                Output.Add(New Tuple(Of List(Of String), Single)(Transcription, PLD1_Transcription.Value.ZipfValue))
            Next
        Next

        'Clearing PLDComparisonCorpus, as this will not be used any more
        PLD1ComparisonCorpus = Nothing

        Return Output

    End Function



    ''' <summary>
    ''' Gets the (zero-based) index of the phone by which the word can be uniquely discriminated from all other words in the comparison corpus
    ''' </summary>
    ''' <returns></returns>
    Public Function GetIsolationPoint(ByVal InputPLD1Transcription As List(Of String), Optional ByRef ErrorString As String = "") As Integer

        Try


            'Prepares an input word transcrptions (which should be a PLD1 type transcription, as in the PLD comparison corpus)
            'Adding the items starting on index 2. Index 0 is the pitch accent/tone, and index 1 is the main stress syllable index
            Dim ModifiedInputTranscription As New List(Of String)
            For PhoneIndex = 2 To InputPLD1Transcription.Count - 1
                ModifiedInputTranscription.Add(InputPLD1Transcription(PhoneIndex))
            Next


            'Go through each phone index and exclude words with different phones at that index, stop when no contrasting words exist

            'Checks if any contrasting words are added, if not 0 is returned since discrimination will be possible at the first phone
            If ComparisonCorpus.Count = 0 Then Return 0

            'Referencing the main ComparisonCorpus as the initial CurrentComparisonCoprus
            Dim CurrentComparisonCoprus As WordGroup.PhoneticComparisonCorpus = ComparisonCorpus

            'Goes through each position in the transcription and stores all words with identical sequences so far in CurrentWordLengthComparisonCorpus
            For i = 0 To ModifiedInputTranscription.Count - 1

                Dim CurrentWordLengthComparisonCorpus As New WordGroup.PhoneticComparisonCorpus

                'Ignores any contrasting spellings that have different or no letters at the current sub-string index
                For Each ComparisonTranscription In CurrentComparisonCoprus

                    'Checks if its the same word, the just skips to next
                    If String.Concat(ModifiedInputTranscription) = String.Concat(ComparisonTranscription.Item1) Then Continue For

                    'Checks if the word end is reached
                    If i >= ComparisonTranscription.Item1.Count Then Continue For

                    'Checks if the current phone is different
                    If ModifiedInputTranscription(i) <> ComparisonTranscription.Item1(i) Then Continue For

                    'Adds the contrasting word (which should have exactly the same word initial phone sequence)
                    CurrentWordLengthComparisonCorpus.Add(ComparisonTranscription)

                Next

                'Replacing the CurrentComparisonCoprus by the CurrentWordLengthComparisonCorpus in order to only compare identical items (so far) in the next step
                CurrentComparisonCoprus = CurrentWordLengthComparisonCorpus

                'Checks if any contrasting words are added, if not i is returned 
                If CurrentComparisonCoprus.Count = 0 Then Return i

            Next

            'Returns the position after the last index, as other words with an initial identical phone sequence but are longer exist
            Return ModifiedInputTranscription.Count

        Catch ex As Exception
            ErrorString &= ex.ToString & vbCrLf
            Return -1
        End Try

    End Function


End Class

