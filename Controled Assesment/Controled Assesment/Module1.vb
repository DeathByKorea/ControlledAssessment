Imports System.Text.RegularExpressions
Imports System.Text

Module Module1

    Sub Main()
        'basic menu
        Console.WriteLine("MENU")
        Console.WriteLine()
        Console.WriteLine("1. Check Password")
        Console.WriteLine("2. Generate Password")
        Console.WriteLine("3. Quit")

        Select Case Console.ReadLine()
            Case "1"
                Check()
            Case "2"
                Generate()
            Case "3"
                Console.WriteLine("Quitting...")
                Threading.Thread.Sleep(250)
                Environment.Exit(0.0)
            Case Else
                Console.WriteLine("Invalid Entry")
                Threading.Thread.Sleep(500)
                Main()
        End Select
    End Sub
    Sub Check()
        Console.Clear()
        'grab the user input
        Console.WriteLine("Enter the Password")
        Dim password As String = Console.ReadLine()
        'test the length
        If Not password.Length >= 8 And password.Length <= 24 Then
            Console.WriteLine("Password is an incorrect length, enter a password between 8 and 24 chars.")
            Threading.Thread.Sleep(1500)
            Main()
        End If
        'check the password contains the correct chars
        If Not CheckChars(password) Then
            Console.WriteLine("Password contains invalid characters or whitespace.")
            Threading.Thread.Sleep(1500)
            Main()
        Else
            Console.WriteLine("Password contains all valid characters!")
        End If
        'get the points for the password
        Dim score As Integer = GetPoints(password)
        'check the strength
        Strength(score)
        Console.ReadLine()
    End Sub
    Sub Generate()
        'get a random string
        Dim password As String = RandomString()
        'get the points of the string
        Dim score As Integer = GetPoints(password)
        If score <= 20 Then
            ' if the password isn't strong, make a new one quietly
            Generate()
        Else
            Strength(score)
        End If

    End Sub
    Function CheckChars(password As String) As Boolean
        Dim match As Boolean = False
        'check it matches every pattern incl. no whitespace
        If Regex.IsMatch(password, "[A-Za-z\d!$%^&*()\-_=+]") And
        Not password.Contains(" ") Then
            Return True
        End If
        Return False
    End Function
    Function GetPoints(password As String) As Integer
        Dim score As Integer = 0
        ' check for individual matches
        score += password.Length
        Console.WriteLine($"Password Length +{password.Length}")
        If Regex.IsMatch(password, "[A-Z]") Then
            score += 5
            Console.WriteLine("Password contains upper case letter +5")
        End If
        If Regex.IsMatch(password, "[a-z]") Then
            Console.WriteLine("Password contains lower case letter +5")
            score += 5
        End If
        If Regex.IsMatch(password, "[\d]") Then
            Console.WriteLine("Password contains digit +5")
            score += 5
        End If
        If Regex.IsMatch(password, "[!$%^&*()\-_=+]") Then
            Console.WriteLine("Password contains symbol +5")
            score += 5
        End If
        ' check for all matches, this probably can be done in a single pattern, but I can not for the life of me 
        ' figure out how in .NET RegEx insead of UNIX
        If Regex.IsMatch(password, "[A-Z]") And
            Regex.IsMatch(password, "[a-z]") And
            Regex.IsMatch(password, "[\d]") And
            Regex.IsMatch(password, "[!$%^&*()\-_=+]") Then
            score += 10
            Console.WriteLine("Password contains all +10")
        End If

        ' Take away points
        If Regex.IsMatch(password, "^[a-zA-Z]*$") Then
            score -= 5
            Console.WriteLine("Password only contains letters -5")
        End If
        If Regex.IsMatch(password, "^[\d]*$") Then
            Console.WriteLine("Password only contains digits -5")
            score -= 5
        End If
        If Regex.IsMatch(password, "^[!$%^&*()\-_=+]*$") Then
            Console.WriteLine("Password only contains symbols -5")
            score -= 5
        End If
        ' RegEx won't match more than one pattern to a substring, so i'm doing it this way.
        Dim patterns As String() = {"qwe", "wer", "ert", "rty", "tyu", "yui", "uio",
                                    "asd", "sdf", "dfg", "fgh", "ghj", "hjk", "jkl",
                                    "zxc", "xcv", "cvb", "vbn", "bnm"}
        For Each pattern As String In patterns
            If password.ToLower.Contains(pattern) Then
                score -= 5
                Console.WriteLine($"Password matches keyboard pattern: {pattern}")
            End If
        Next
        Return score
    End Function

    Sub Strength(points As Integer)
        'check the strength of the strings
        If points >= 20 Then
            Console.WriteLine("You have a strong password!")
            Console.WriteLine($"Your password scored {points} Points!")
        ElseIf points <= 0 Then
            Console.WriteLine("You have a weak password!")
            Console.WriteLine($"Your password scored {points} Points!")
        Else
            Console.WriteLine($"Your password scored {points} Points!")
        End If

    End Sub
    Function RandomString() As String
        'make a random string using randoms and stringbuilders
        Dim s As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuvwxyz!$%^&*()-_=+"
        Static r As New Random
        'between 8 and 12 inclusive
        Dim chactersInString As Integer = r.Next(8, 12)
        Dim sb As New StringBuilder
        For i As Integer = 1 To chactersInString
            Dim idx As Integer = r.Next(0, s.Length)
            sb.Append(s.Substring(idx, 1))
        Next
        Return sb.ToString()
    End Function
End Module
