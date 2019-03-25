Imports System.IO

Module Module1

    Sub Main()

        Dim timeStamp As String = DateTime.Now.ToString("yyyyMMddHHmmssfff")

        Dim filePath As String = Path.Combine(Path.GetTempPath(), timeStamp & "-test.csv")
        Console.WriteLine(filePath)

        Dim logFile As String = Path.Combine(Path.GetTempPath(), timeStamp & "-log.txt")
        LogToFile(logFile, "File path: " & filePath)

        CreateCSVFile(filePath)

    End Sub

    Sub CreateCSVFile(ByVal filePath As String)
        Dim stringBuildOutput As New Text.StringBuilder()

        stringBuildOutput.AppendLine("Date, RecipeName, RecipePath, StartTime, CompleteTime, Result")


        stringBuildOutput.AppendLine("20190301, Test1, C:\user1\path1\fil1.csv, 20190301 11：00, 20190301 11:05, Success")
        stringBuildOutput.AppendLine("20190302, Test2, C:\user2\path2\fil2.csv, 20190301 12：00, 20190301 12:05, Fail")
        stringBuildOutput.AppendLine("20190303, Test1, C:\user3\path3\fil3.csv, 20190301 13：00, 20190301 13:05, Success")

        File.WriteAllText(filePath, stringBuildOutput.ToString())
    End Sub

    Sub LogToFile(ByVal logFile As String, ByVal message As String)
        Using file As StreamWriter = System.IO.File.AppendText(logFile)
            file.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:sss") & "   " & message)
        End Using
    End Sub

End Module
