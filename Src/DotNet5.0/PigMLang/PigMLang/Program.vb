Imports System.IO
Module Program
    Sub Main(args As String())
        Console.Write("Use the default culture for initialization...")
        Dim oPigMLang As New PigMLang()
        If oPigMLang.LastErr <> "" Then
            Console.WriteLine(oPigMLang.LastErr)
        Else
            Console.WriteLine("OK")
        End If
        Console.Write("LoadMLangInf...")
        oPigMLang.LoadMLangInf()
        If oPigMLang.LastErr <> "" Then
            Console.WriteLine(oPigMLang.LastErr)
        Else
            Console.WriteLine(oPigMLang.GetMLangText("OK", "OK"))
            Console.WriteLine(oPigMLang.GetMLangText("AppTitle", "AppTitle: ") & oPigMLang.AppTitle)
            Console.WriteLine(oPigMLang.GetMLangText("AppPath", "AppPath: ") & oPigMLang.AppPath)

            Do While True
                Console.WriteLine(oPigMLang.GetMLangText("CurrCultureName", "CurrCultureName: ") & oPigMLang.CurrCultureName)
                Console.WriteLine(oPigMLang.GetMLangText("CurrLCID", "CurrLCID: ") & oPigMLang.CurrLCID.ToString)
                Console.WriteLine(oPigMLang.GetMLangText("CurrMLangTitle", "CurrMLangTitle: ") & oPigMLang.CurrMLangTitle)
                Console.WriteLine(oPigMLang.GetMLangText("CurrMLangFile", "CurrMLangFile: ") & oPigMLang.CurrMLangFile)
                Console.WriteLine(oPigMLang.GetMLangText("PressQuitDesc", "Press the Escape (Esc) key to quit."))
                Console.WriteLine(oPigMLang.GetMLangText("PressGetAllLangInfTab", "Press A to GetAllLangInf(TAB)"))
                Console.WriteLine(oPigMLang.GetMLangText("PressGetAllLangInfMD", "Press B to GetAllLangInf(Markdown)"))
                Console.WriteLine(oPigMLang.GetMLangText("PressSetCurrLang", "Press C to Set Current Culture"))
                Select Case Console.ReadKey().Key
                    Case ConsoleKey.Escape
                        Exit Do
                    Case ConsoleKey.A
                        Console.WriteLine(oPigMLang.GetAllLangInf(PigMLang.enmGetInfFmt.TabSeparator))
                    Case ConsoleKey.B
                        Console.WriteLine(oPigMLang.GetAllLangInf(PigMLang.enmGetInfFmt.Markdown))
                    Case ConsoleKey.C
                        Console.Write(oPigMLang.GetMLangText("RefCanUseCultureList", "Refresh Can Use Culture List..."))
                        oPigMLang.RefCanUseCultureList()
                        If oPigMLang.LastErr <> "" Then
                            Console.WriteLine(oPigMLang.LastErr)
                        Else
                            Console.WriteLine(oPigMLang.GetMLangText("OK", "OK"))
                            Dim intMax As Integer = oPigMLang.CanUseCultureList.Count - 1
                            For i = 0 To intMax
                                With oPigMLang.CanUseCultureList(i)
                                    Console.WriteLine("[{0}]{1}", i, .DisplayName)
                                End With
                            Next
                            Do While True
                                Console.Write(oPigMLang.GetMLangText("MLang", "SeleCurrCulture", "Input number to select current Culture: "))
                                Dim chrSelect As Char = Console.ReadKey().KeyChar
                                Select Case chrSelect
                                    Case "0" To intMax.ToString
                                        Console.WriteLine("...")
                                        Dim intItem As Integer = CInt(chrSelect.ToString)
                                        Dim intLCID As Integer = oPigMLang.CanUseCultureList(intItem).LCID
                                        Console.Write("SetCurrCulture...")
                                        oPigMLang.SetCurrCulture(intLCID)
                                        If oPigMLang.LastErr <> "" Then
                                            Console.WriteLine(oPigMLang.LastErr)
                                        Else
                                            Console.WriteLine(oPigMLang.GetMLangText("OK", "OK"))
                                            Console.Write(oPigMLang.GetMLangText("MLang", "ReLoadMLangInf", "ReLoadMLangInf..."))
                                            oPigMLang.LoadMLangInf()
                                            If oPigMLang.LastErr <> "" Then
                                                Console.WriteLine(oPigMLang.LastErr)
                                            Else
                                                Console.WriteLine(oPigMLang.GetMLangText("OK", "OK"))
                                            End If
                                        End If
                                        Exit Do
                                    Case Else
                                        Console.WriteLine(oPigMLang.GetMLangText("MLang", "PleaseInp0To", "Please input 0 to ") & intMax.ToString)
                                End Select

                            Loop
                        End If
                End Select
            Loop
        End If

    End Sub
End Module
