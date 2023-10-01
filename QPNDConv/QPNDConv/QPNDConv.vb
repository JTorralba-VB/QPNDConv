Imports System
Imports System.IO

Module QPNDConv

    Sub Main()

        Dim Parameters() As String

        If Command.Length <= 0 Then
            MsgBox("Invalid Parameters")
            Exit Sub
        Else
            Parameters = Command.Split("*")
            Console.Clear()
            'Dim Parameter As String
            'For Each Parameter In Parameters
            'Console.WriteLine(Parameter)
            'Next
        End If


        Dim INPUT1 As String
        Dim INPUT2 As String
        Dim OUTPUT As String
        Dim TROUBLESHOOT As String

        INPUT1 = Parameters(0)
        INPUT2 = Parameters(1)
        OUTPUT = Parameters(2)
        TROUBLESHOOT = Parameters(3)

        Dim SubTotal1 As Integer
        Dim SubTotal2 As Integer

        Dim OriginalContents As String
        Dim ModifiedContents As String

        Console.Write(vbCrLf)
        Console.Write(vbCrLf)

        Console.WriteLine("SPLITTING RECORDS SOURCE #1 --------------------------------")
        OriginalContents = FileRead(INPUT1)
        ModifiedContents = FileModify(OriginalContents)
        FileWrite("C:\QPND_ONE.TXT", ModifiedContents)

        Console.Write(vbCrLf)
        Console.Write(vbCrLf)

        Console.WriteLine("SPLITTING RECORDS SOURCE #2 --------------------------------")
        OriginalContents = FileRead(INPUT2)
        ModifiedContents = FileModify(OriginalContents)
        FileWrite("C:\QPND_TWO.TXT", ModifiedContents)

        Console.Write(vbCrLf)
        Console.Write(vbCrLf)

        DelFile(OUTPUT)

        Console.WriteLine("EXTRACTING DATA SOURCE #1 ----------------------------------")
        SubTotal1 = FileWriteExtract("C:\QPND_ONE.TXT", OUTPUT, TROUBLESHOOT)

        Console.WriteLine("EXTRACTING DATA SOURCE #2 ----------------------------------")
        SubTotal2 = FileWriteExtract("C:\QPND_TWO.TXT", OUTPUT, TROUBLESHOOT)

        Console.WriteLine("SOURCE1     " + Str(SubTotal1))
        Console.WriteLine("SOURCE2     " + Str(SubTotal2))
        Console.WriteLine("============================================================")
        Console.WriteLine("TOTAL       " + Str(SubTotal1 + SubTotal2))

        Console.Write(vbCrLf)
        Console.Write(vbCrLf)

        'Console.Write("PRESS <ENTER> TO EXIT")
        'Console.Read()

    End Sub

    Private Sub DelFile(ByVal PhysicalFile As String)
        Dim FileExist As String
        FileExist = Dir$(PhysicalFile)
        If FileExist = "QPND_CLN.TXT" Then
            Kill(PhysicalFile)
        End If
    End Sub

    Function FileRead(ByVal InputFile As String) As String
        Dim FileStreamRead As StreamReader
        Dim Contents As String
        FileStreamRead = New StreamReader(InputFile)
        Contents = FileStreamRead.ReadToEnd()
        FileStreamRead.Close()
        Return Contents
    End Function

    Private Sub FileWrite(ByVal OutputFile As String, ByVal Contents As String)
        Dim FileStreamWriter As StreamWriter
        FileStreamWriter = New StreamWriter(OutputFile)
        FileStreamWriter.Write(Contents)
        FileStreamWriter.Close()
    End Sub

    Function FileModify(ByVal Contents As String) As String
        Dim Contents2 As String
        Contents2 = Replace(Contents, "   * ", vbCrLf)
        Return Contents2
    End Function

    Function FileWriteExtract(ByVal InputFile As String, ByVal OUTPUT As String, ByVal TROUBLESHOOT As String) As Integer

        Dim FileStreamRead As StreamReader
        Dim FileStreamWriter As StreamWriter

        Dim Customer_Phone As String
        Dim Customer_Address As String
        Dim Street_Number As String
        Dim Street_Name As String
        Dim Customer_MiscAddress As String
        Dim Customer_Name As String
        Dim Customer_TelCo As String
        Dim Customer_City As String

        Dim Comment As String

        Dim Record As String
        Dim Counter As Integer

        Counter = 0

        FileStreamRead = New StreamReader(InputFile)
        FileStreamWriter = New StreamWriter(OUTPUT, True)

        While Not (FileStreamRead.EndOfStream)

            Comment = Nothing
            Record = FileStreamRead.ReadLine()

            Select Case Trim(Left(Record, 3))

                Case Is = "UHL"
                    'MsgBox("UHL Header")
                Case Is = "UTL"
                    'MsgBox("UTL Header")
                Case Is = ""
                    'MsgBox("Empty")
                Case Else

                    Counter = Counter + 1

                    Customer_Phone = Replace(Trim(Record.Substring(0, 10)), "  ", " ")
                    Street_Number = Replace(Trim(Record.Substring(10, 16)), "  ", " ")
                    Street_Name = Replace(Trim(Record.Substring(26, 66)), "  ", " ")
                    Customer_Address = Replace(Trim(Trim(Street_Number) + " " + Trim(Street_Name)), "  ", " ")
                    Customer_MiscAddress = Replace(Trim(Record.Substring(126, 50)), "  ", " ")
                    Customer_Name = Replace(Trim(Record.Substring(186, 32)), "  ", " ")
                    Customer_TelCo = Replace(Trim(Record.Substring(259, 5)), "  ", " ")
                    Customer_City = Replace(Trim(Record.Substring(92, 32)), "  ", " ")

                    If Customer_Phone.Length <> 10 Then
                        Comment = Comment + "invalid 10 digit phone"
                    End If

                    If Street_Number = "" Then
                        Comment = Comment + "invalid street number"
                    End If

                    If Street_Name = "" Then
                        Comment = Comment + "invalid street name"
                    End If

                    If Customer_Name = "" Then
                        Comment = Comment + "invalid customer name"
                    End If

                    If TROUBLESHOOT = "DEBUG" Then
                        Console.WriteLine(Trim(Str(Counter)) + "     " + Customer_Phone + "     " + Customer_Address)
                    End If

                    If Comment <> "" Then
                        'MsgBox(Comment)
                    End If

                    If Street_Name = "SUMMER" Then
                        MsgBox(Street_Name)
                    End If

                    FileStreamWriter.Write(Street_Name + vbTab + Street_Number + vbTab + Customer_Phone + vbTab + Customer_Address + vbTab + Customer_MiscAddress + vbTab + Customer_Name + vbTab + Customer_TelCo + vbTab + Comment + vbTab + Customer_City)
                    'FileStreamWriter.Write(Customer_Phone + vbTab + Customer_City)
                    FileStreamWriter.Write(vbCrLf)

            End Select

        End While

        FileStreamWriter.Close()
        FileStreamRead.Close()

        Kill(InputFile)

        Console.Write(vbCrLf)
        Console.Write(vbCrLf)

        Return Counter
    End Function

End Module
