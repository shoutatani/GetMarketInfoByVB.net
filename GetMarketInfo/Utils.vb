Module Utils

    Public Sub LogStart(ByVal argsMessage As String)
        Console.WriteLine(argsMessage & " - Start")
    End Sub

    Public Sub LogEnd(ByVal argsMessage As String)
        Console.WriteLine(argsMessage & " - End")
    End Sub

    Public Sub Logging(ByVal argsMessage As String)
        Console.WriteLine("[Info]" & argsMessage)
    End Sub

End Module
