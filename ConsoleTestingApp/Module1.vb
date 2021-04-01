Module Module1

    Sub Main()

        'Creates a new instance of WordMetricsCalculation
        Dim MWC As New SWMC.WordMetricsCalculation()

        'Creates an array of (lower case) words for which to calculate word metrics
        Dim InputWords As String() = {"hej", "du"}

        'Calculates the word metrics of the input words
        Dim Output = MWC.CalculateWordMetrics(InputWords)

        'Sets up a column order for the output data
        Dim ColumnOrder = New SWMC.WordListsIO.PhoneticTxtStringColumnIndices
        ColumnOrder.SetWebSiteColumnOrder(False, True, True, True, True, True, True, True, True)

        'Displays the result in the console
        Console.OutputEncoding = System.Text.Encoding.UTF8
        Console.WriteLine(ColumnOrder.GetColumnHeadingsString())
        For Each entry In Output.MemberWords
            Console.WriteLine(entry.GenerateFullPhoneticOutputTxtString(ColumnOrder))
        Next

        'Hold the colsone open so the results can be seen
        Console.Read()

    End Sub

End Module
