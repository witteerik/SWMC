Imports System

Module Program
    Sub Main(args As String())

        'Creates a new instance of WordMetricsCalculation
        Dim MWC As New SWMC.WordMetricsCalculation()

        'Creates an array of (lower case) words for which to calculate word metrics
        Dim InputWords As String() = {"hej", "du"}

        Dim Hits As Integer = 0
        Dim testLookUp = SWMC.AfcListSearch.SearchAfcList("SELECT * FROM AfcList WHERE OrthographicForm LIKE 'hej';", Hits)

        'Calculates the word metrics of the input words
        Dim Output = MWC.CalculateWordMetrics(InputWords)

        Dim ntow = Output.MemberWords(0).ConvertToTextOnlyWord()

        Dim ErrorList As New List(Of String)
        Dim bcw = ntow.ConvertToWord(ErrorList)

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
