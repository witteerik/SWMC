Imports System

Module Program
    Sub Main(args As String())

        'Try
        '    System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("sv-SE")
        '    'System.Threading.Thread.CurrentThread.CurrentUICulture = System.Globalization.CultureInfo.CreateSpecificCulture("sv-SE")
        '    'System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture
        'Catch ex As Exception
        '    Console.WriteLine("Warning: Unavailable language pack. The local computer may lack some required language components! You can continue running the program but it may not function correctly!")
        'End Try
        ''Setting the negative sign to hyphen
        'Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NegativeSign = "-"

        'Creates a new instance of WordMetricsCalculation
        Dim MWC As New SWMC.WordMetricsCalculation()

        'Creates an array of (lower case) words for which to calculate word metrics
        Dim InputWords As String() = {"hej", "du"}

        Dim Hits As Integer = 0
        Dim testLookUp = SWMC.AfcListSearch.SearchAfcList("SELECT * FROM AfcList WHERE OrthographicForm LIKE 'hej';", Hits)

        Console.WriteLine("Calculatiing word metrics for" & Hits & " words.")

        'Calculates the word metrics of the input words
        Dim ErrorString As String = ""
        Dim Output = MWC.CalculateWordMetrics(InputWords,,,,,,,,,,,,,,,,,,,,, ErrorString)

        If ErrorString.Trim <> "" Then Console.WriteLine(ErrorString)

        If Output Is Nothing Then
            Console.WriteLine("No Wordgroup returned!")
            Console.WriteLine("Press any key to continue")
            Console.ReadLine()
            Exit Sub
        End If

        If Output.MemberWords.Count = 0 Then
            Console.WriteLine("No words in Wordgroup!")
            Console.WriteLine("Press any key to continue")
            Console.ReadLine()
        End If

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
