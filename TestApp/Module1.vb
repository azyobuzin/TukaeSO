Module Module1
    Sub Main()
        'Application.EnableVisualStyles()

        Dim r As New Random

        Do
            Console.WriteLine(CreatePassword(r.Next(1, 20)))
            Threading.Thread.Sleep(1000)
        Loop

        'Console.ReadKey(True)
    End Sub
End Module
