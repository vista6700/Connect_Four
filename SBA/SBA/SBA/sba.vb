Module sba
    Dim B(3), i, j As Integer
    Dim Player_Name(2) As String
    Dim Number(1) As Char
    Dim ewcheckerboard(5, 6) As String
    Dim startevent, endevent As Date
    Sub Main()
        Console.WriteLine("遊戲 : 「四子棋」")
        Console.WriteLine("------------------------------")
        Console.WriteLine("(1) 二人對奕")
        Console.WriteLine("(2) 人對電腦(沒有人工智能)")
        Console.WriteLine("(3) 排名榜")
        Console.WriteLine("(4) 離開遊戲")
        Console.WriteLine("------------------------------")
        For i = 0 To 3
            B(i) = 0
        Next
        For i = 0 To 5
            For j = 0 To 6
                ewcheckerboard(i, j) = "○"
            Next
        Next
        Do
            Console.Write("遊戲模式: ")
            Number(1) = Console.ReadKey.KeyChar
            Console.WriteLine()
        Loop Until Number(1) = "1" Or Number(1) = "2" Or Number(1) = "3" Or Number(1) = "4"
        Select Case Val(Number(1))
            Case 1
                Game_Mode1()
            Case 2
                Game_Mode2()
            Case 3
                Game_Mode4()
            Case 4
                End
        End Select
    End Sub
    Sub Game_Mode1()
        Console.Write("用戶名稱 (1) 輸入後按下Enter以繼續: ")
        Player_Name(1) = Console.ReadLine()
        Console.Write("用戶名稱 (2) 輸入後按下Enter以繼續: ")
        Player_Name(2) = Console.ReadLine()
        While Player_Name(1) = Player_Name(2)
            Console.WriteLine("請選擇並更改用戶名稱")
            Console.WriteLine("用戶名稱 (1): [{0}] or 用戶名稱 (2): [{1}]", Player_Name(1), Player_Name(2))
            Do
                Number(0) = Console.ReadKey.KeyChar
                Console.WriteLine()
            Loop Until Number(0) = "1" Or Number(0) = "2"
            If Number(0) = "1" Then
                Console.Write("請輸入用戶名稱 (1) 輸入後按下Enter以繼續: ")
                Player_Name(1) = Console.ReadLine()
            Else
                Console.Write("請輸入用戶名稱 (2) 輸入後按下Enter以繼續: ")
                Player_Name(2) = Console.ReadLine
            End If
        End While
        startevent = TimeOfDay
        Do
            B(0) = 1
            Player_Enter()
            Newcheckerboard()
            B(0) = 2
            Player_Enter()
            Newcheckerboard()
        Loop
    End Sub
    Sub Game_Mode2()
        Console.Write("用戶名稱 {輸入後按下Enter以繼續}: ")
        Player_Name(1) = Console.ReadLine()
        startevent = TimeOfDay
        Do
            B(1) = 1
            Player_Enter()
            Newcheckerboard()
            Number(0) = Chr(Asc(Str(Int(Rnd() * 7) + 1).Trim))
            Player_Name(2) = "GameMode_2"
            B(1) = 2
            Newcheckerboard()
        Loop
    End Sub
    Sub Game_Mode4()
        Dim sw As IO.StreamReader = IO.File.OpenText("Game_Mode4.txt")
        Dim sw2 As IO.StreamReader = IO.File.OpenText("Game_Mode2.txt")
        Dim OPENSW As String
        Console.WriteLine("(1)人對電腦(沒有人工智能)排行榜")
        Console.WriteLine("(2)二人對奕排行榜")
        Number(1) = Console.ReadKey.KeyChar
        Console.WriteLine()
        While Val(Number(1)) < 1 Or Val(Number(1)) > 2
            Console.Write("請再次輸入排行榜")
            Number(1) = Console.ReadKey.KeyChar
            Console.WriteLine()
        End While
        If Number(1) = "2" Then
            OPENSW = sw.ReadLine()
            While OPENSW <> Nothing
                Console.WriteLine(OPENSW)
                OPENSW = sw.ReadLine()
            End While
        End If
        If Number(1) = "1" Then
            OPENSW = sw2.ReadLine()
            While OPENSW <> Nothing
                Console.WriteLine(OPENSW)
                OPENSW = sw2.ReadLine()
            End While
        End If
        sw2.Close()
        sw.Close()
        Main()
    End Sub
    Sub file()
        endevent = TimeOfDay
        Console.WriteLine("時間（秒）:{0}", time_count(startevent, endevent))
        Dim outfile As IO.StreamWriter = IO.File.AppendText("Game_Mode4.txt")
        Dim outfile1 As IO.StreamWriter = IO.File.AppendText("Game_Mode2.txt")
        If B(0) = 3 Or B(0) = 4 Then
            outfile.WriteLine("時間（秒）:{0}", time_count(startevent, endevent))
            outfile.WriteLine(Player_Name(1) & "vs" & Player_Name(2))
            If B(0) = 3 Then
                outfile.WriteLine(Player_Name(1) & "已勝出此遊戲")
            ElseIf B(0) = 4 Then
                outfile.WriteLine(Player_Name(2) & "已勝出此遊戲")
            End If
            If B(0) = 6 Then
                outfile.WriteLine(Player_Name(1) & "(平局)" & Player_Name(2))
            End If
            outfile.WriteLine("===")
        ElseIf B(1) = 1 Or B(1) = 2 Then
            outfile1.WriteLine("時間（秒）:{0}", time_count(startevent, endevent))
            outfile1.WriteLine(Player_Name(1) & "vs Game_Mode2")
            If B(1) = 3 Then
                outfile1.WriteLine("Game_Mode2,已勝出此遊戲")
            ElseIf B(2) = 7 Then
                outfile1.WriteLine(Player_Name(1) & ",已勝出此遊戲")
            End If
            If B(2) = 5 Then
                outfile1.WriteLine(Player_Name(1) & "(平局)Game_Mode(2)")
            End If
            outfile1.WriteLine("===")
        End If
        outfile.Close()
        outfile1.Close()
    End Sub
    Sub Newcheckerboard()
        i = 0
        Do
            If ewcheckerboard(i, Val(Number(0)) - 1) <> "○" Then
                Console.WriteLine("錯誤")
                If B(0) = 1 Or B(1) = 1 Or B(2) = 1 Then
                    Player_Enter()
                ElseIf B(0) = 2 Then
                    Player_Enter()
                ElseIf B(1) = 2 Then
                    Number(0) = Chr(Asc(Str(Int(Rnd() * 7) + 1).Trim))
                End If
            End If
        Loop Until ewcheckerboard(i, Val(Number(0)) - 1) = "○"
        i = 5
        While ewcheckerboard(i, Val(Number(0)) - 1) <> "○"
            i -= 1
        End While
        If B(0) = 1 Or B(1) = 1 Or B(3) = 1 Then
            ewcheckerboard(i, Val(Number(0)) - 1) = "●"
        ElseIf B(0) = 2 Or B(1) = 2 Or B(3) = 2 Then
            ewcheckerboard(i, Val(Number(0)) - 1) = "■"
        End If
        Console.WriteLine("1.        2.        3.        4.        5.        6.        7.")
        Console.WriteLine()
        For i = 0 To 5
            For j = 0 To 6
                Console.Write(ewcheckerboard(i, j) & "        ")
            Next
            Console.WriteLine()
            Console.WriteLine()
        Next
        Console.WriteLine("------------------------------")
        For i = 0 To 2
            For j = 0 To 3
                If ewcheckerboard(i, j) = "●" And ewcheckerboard(i + 1, j + 1) = "●" And ewcheckerboard(i + 2, j + 2) = "●" And ewcheckerboard(i + 3, j + 3) = "●" Then
                    Console.WriteLine("{0} ,已勝出此遊戲", Player_Name(1))
                    B(0) = 3
                    B(2) = 7
                ElseIf ewcheckerboard(i, j) = "■" And ewcheckerboard(i + 1, j + 1) = "■" And ewcheckerboard(i + 2, j + 2) = "■" And ewcheckerboard(i + 3, j + 3) = "■" Then
                    Console.WriteLine("{0} ,已勝出此遊戲", Player_Name(2))
                    If B(1) = 2 Or B(2) = 2 Then
                        B(1) = 3
                    ElseIf B(0) = 2 Then
                        B(0) = 4
                    End If
                ElseIf ewcheckerboard(i + 3, j) = "●" And ewcheckerboard(i + 2, j + 1) = "●" And ewcheckerboard(i + 1, j + 2) = "●" And ewcheckerboard(i, j + 3) = "●" Then
                    Console.WriteLine("{0} ,已勝出此遊戲", Player_Name(1))
                    B(0) = 3
                    B(2) = 7
                ElseIf ewcheckerboard(i + 3, j) = "■" And ewcheckerboard(i + 2, j + 1) = "■" And ewcheckerboard(i + 1, j + 2) = "■" And ewcheckerboard(i, j + 3) = "■" Then
                    Console.WriteLine("{0} ,已勝出此遊戲", Player_Name(2))
                    If B(1) = 2 Or B(2) = 2 Then
                        B(1) = 3
                    ElseIf B(0) = 2 Then
                        B(0) = 4
                    End If
                End If
            Next
            For j = 0 To 6
                If ewcheckerboard(i + 3, j) = "●" And ewcheckerboard(i + 2, j) = "●" And ewcheckerboard(i + 1, j) = "●" And ewcheckerboard(i, j) = "●" Then
                    Console.WriteLine("{0} ,已勝出此遊戲", Player_Name(1))
                    B(0) = 3
                    B(2) = 7
                ElseIf ewcheckerboard(i + 3, j) = "■" And ewcheckerboard(i + 2, j) = "■" And ewcheckerboard(i + 1, j) = "■" And ewcheckerboard(i, j) = "■" Then
                    Console.WriteLine("{0} ,已勝出此遊戲", Player_Name(2))
                    If B(1) = 2 Or B(2) = 2 Then
                        B(1) = 3
                    ElseIf B(0) = 2 Then
                        B(0) = 4
                    End If
                End If
            Next
        Next
        For i = 0 To 5
            For j = 0 To 3
                If ewcheckerboard(i, j) = "●" And ewcheckerboard(i, j + 1) = "●" And ewcheckerboard(i, j + 2) = "●" And ewcheckerboard(i, j + 3) = "●" Then
                    Console.WriteLine("{0} ,已勝出此遊戲", Player_Name(1))
                    B(0) = 3
                    B(2) = 7
                ElseIf ewcheckerboard(i, j) = "■" And ewcheckerboard(i, j + 1) = "■" And ewcheckerboard(i, j + 2) = "■" And ewcheckerboard(i, j + 3) = "■" Then
                    Console.WriteLine("{0} ,已勝出此遊戲", Player_Name(2))
                    If B(1) = 2 Or B(2) = 2 Then
                        B(1) = 3
                    ElseIf B(0) = 2 Then
                        B(0) = 4
                    End If
                End If
            Next
        Next
        If ewcheckerboard(0, 0) <> "○" And ewcheckerboard(0, 1) <> "○" And ewcheckerboard(0, 2) <> "○" And ewcheckerboard(0, 3) <> "○" And ewcheckerboard(0, 4) <> "○" And ewcheckerboard(0, 5) <> "○" And ewcheckerboard(0, 6) <> "○" Then
            Console.WriteLine("平局")
            B(0) = 6
            B(2) = 5
        End If
        If B(0) = 3 Or B(0) = 4 Or B(0) = 6 Or B(1) = 3 Or B(2) = 5 Or B(2) = 7 Then
            file()
            Main()
        End If
    End Sub
    Function time_count(ByVal sTime As Date, ByVal eTime As Date)
        Dim sHour, sMin, sSec As Integer
        Dim eHour, eMin, eSec As Integer
        sHour = sTime.Hour
        sMin = sTime.Minute
        sSec = sTime.Second
        eHour = eTime.Hour
        eMin = eTime.Minute
        eSec = eTime.Second
        Return (eHour - sHour) * 3600 + (eMin - sMin) * 60 + (eSec - sSec)
    End Function
    Sub Player_Enter()
        If B(0) = 1 Or B(1) = 1 Or B(3) = 1 Then
            Console.Write("{0}投入棋盤: ", Player_Name(1))
        ElseIf B(0) = 2 Then
            Console.Write("{0}投入棋盤: ", Player_Name(2))
        End If
        Number(0) = Console.ReadKey.KeyChar
        Console.WriteLine()
        If Number(0) = "x" Or Number(0) = "X" Then
            End
        End If
        While Val(Number(0)) > 7 Or Val(Number(0)) < 1
            If B(0) = 1 Or B(1) = 1 Then
                Console.Write("請{0}再次投入棋盤: ", Player_Name(1))
            ElseIf B(0) = 2 Then
                Console.Write("請{0}再次投入棋盤: ", Player_Name(2))
            End If
            Number(0) = Console.ReadKey.KeyChar
            Console.WriteLine()
            If Number(0) = "x" Or Number(0) = "X" Then
                End
            End If
        End While
    End Sub
End Module