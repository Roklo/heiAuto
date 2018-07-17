Imports Microsoft.Office.Interop.Excel
Module modTagGeneration


    Dim excelSheet As Worksheet
    Dim xlWorkbook = Nothing
    Private Excel As Application = New Application()

    Private Sub LoadIOLIst()


        Dim xlWorkbook = Excel.Workbooks.Open("")
        xlWorkbook = findWorkBook(modConfiguration.IOList_Path)


        If xlWorkbook Is Nothing Then
            xlWorkbook = Excel.Workbooks.Open(modConfiguration.IOList_Path, , True)

        End If

        excelSheet = xlWorkbook.Sheets(modConfiguration.IOLIST_Sheet)

    End Sub

    Function CalcAddress()

        LoadIOLIst()
        Dim X = 9
        Dim a = 1
        Dim b = 1
        Dim c = 1
        Dim d = 1

        While excelSheet.Cells(X, "A") <> ""

            ' find and asign tags for alarms

            If excelSheet.Cells(X, NODE) <> "" Then
                If Right(excelSheet.Cells(X, SN), 4) = "Port" Or Right(excelSheet.Cells(X, SN), 4) = "Stbd" Then
                    excelSheet.Cells(X, Adr2) = CalcReg(excelSheet.Cells(X, SN), excelSheet.Cells(X, IOCNO), excelSheet.Cells(X, CH))

                Else
                    excelSheet.Cells(X, Adr2) = excelSheet.Cells(X, MbAdd)
                End If
                a = a + 1

            End If


            X = X + 1
        End While


    End Function

    Function Alarm()

        LoadIOLIst()

        Dim X = 9
        Dim a = SANP
        Dim b = SANS
        Dim c = 1
        Dim d = 1

        While excelSheet.Cells(X, "A") <> ""

            ' find and asign tags for alarms

            If (UCase(excelSheet.Cells(X, AYN)) = "YES" Or UCase(excelSheet.Cells(X, AYN)) = "MON") And (Right(excelSheet.Cells(X, NODE), 4) = "Port" Or Right(excelSheet.Cells(X, NODE), 5) = "Cab01") Then
                excelSheet.Cells(X, IasTN) = "Alm" & Format(a, "#000#")
                If Right(excelSheet.Cells(X, SN), 4) = "Port" Or Right(excelSheet.Cells(X, SN), 4) = "Stbd" Then
                    excelSheet.Cells(X, Adr) = CalcReg(excelSheet.Cells(X, SN), excelSheet.Cells(X, IOCNO), excelSheet.Cells(X, CH))
                    Excel.Sheets("IO List").Select
                Else
                    excelSheet.Cells(X, Adr) = excelSheet.Cells(X, MbAdd)
                End If
                a = a + 1
            ElseIf (UCase(excelSheet.Cells(X, AYN)) = "YES" Or UCase(excelSheet.Cells(X, AYN)) = "MON") And (Right(excelSheet.Cells(X, NODE), 4) = "Stbd" Or Right(excelSheet.Cells(X, NODE), 5) = "Cab02") Then
                excelSheet.Cells(X, IasTN) = "Alm" & (Val(b) + MaxAlm)
                If Right(excelSheet.Cells(X, SN), 4) = "Port" Or Right(excelSheet.Cells(X, SN), 4) = "Stbd" Then
                    excelSheet.Cells(X, Adr) = CalcReg(excelSheet.Cells(X, SN), excelSheet.Cells(X, IOCNO), excelSheet.Cells(X, CH))
                    Excel.Sheets("IO List").Select
                Else
                    excelSheet.Cells(X, Adr) = excelSheet.Cells(X, MbAdd)
                End If
                b = b + 1
            End If


            X = X + 1
        End While


    End Function

    Function CalcReg(CabNo As String, ModNo As String, ChNo As Integer) As String

        LoadIOLIst()

        Dim lRegAddress As Worksheet
        lRegAddress = Excel.Sheets("RegAddress")
        'Sheets("RegAddress").Select

        Dim i = Excel.i = 4
        For i = 4 To 45
            If UCase(lRegAddress.Cells(i, "A")) = ModNo Then

                Select Case UCase(CabNo)

                    Case "CAB01PORT"
                        CalcReg = lRegAddress.Cells(i, "C") & ":" & (lRegAddress.Cells(i, "D") + ChNo - 1)
                    Case "CAB02STBD"
                        CalcReg = lRegAddress.Cells(i, "C") & ":" & (lRegAddress.Cells(i, "E") + ChNo - 1)
                    Case "CAB03PORT"
                        CalcReg = lRegAddress.Cells(i, "C") & ":" & (lRegAddress.Cells(i, "F") + ChNo - 1)
                    Case "CAB04STBD"
                        CalcReg = lRegAddress.Cells(i, "C") & ":" & (lRegAddress.Cells(i, "G") + ChNo - 1)
                    Case "CAB05PORT"
                        CalcReg = lRegAddress.Cells(i, "C") & ":" & (lRegAddress.Cells(i, "H") + ChNo - 1)
                    Case "CAB06STBD"
                        CalcReg = lRegAddress.Cells(i, "C") & ":" & (lRegAddress.Cells(i, "I") + ChNo - 1)
                    Case "CAB07PORT"
                        CalcReg = lRegAddress.Cells(i, "C") & ":" & (lRegAddress.Cells(i, "J") + ChNo - 1)
                    Case "CAB08STBD"
                        CalcReg = lRegAddress.Cells(i, "C") & ":" & (lRegAddress.Cells(i, "K") + ChNo - 1)
                    Case "CAB09PORT"
                        CalcReg = lRegAddress.Cells(i, "C") & ":" & (lRegAddress.Cells(i, "L") + ChNo - 1)
                    Case "CAB10STBD"
                        CalcReg = lRegAddress.Cells(i, "C") & ":" & (lRegAddress.Cells(i, "M") + ChNo - 1)
                    Case "CAB11PORT"
                        CalcReg = lRegAddress.Cells(i, "C") & ":" & (lRegAddress.Cells(i, "N") + ChNo - 1)
                    Case "CAB12STBD"
                        CalcReg = lRegAddress.Cells(i, "C") & ":" & (lRegAddress.Cells(i, "O") + ChNo - 1)

                End Select

                GoTo line1
            End If
        Next i

line1:
    End Function

    Function valve()

        LoadIOLIst()


        Dim X = 9
        Dim a = 1
        Dim b = 1
        Dim c = 1
        Dim d = 1
        Dim e = 2

        While excelSheet.Cells(X, "A") <> ""

            'find and asign tags for valves
            If X = 385 Then
                X = 385
            End If
            If (UCase(excelSheet.Cells(X, ISA)) = "OPN" Or UCase(excelSheet.Cells(X, ISA)) = "CLS" Or UCase(excelSheet.Cells(X, ISA)) = "FOP" Or UCase(excelSheet.Cells(X, ISA)) = "FCL") And UCase(Right(excelSheet.Cells(X, NODE), 4)) = "PORT" Then
                excelSheet.Cells(X, IasTN) = "RCV" & c & "_" & UCase(excelSheet.Cells(X, ISA))

                If excelSheet.Cells(X, IOCNO) <> "" Then
                    excelSheet.Cells(X, Adr) = CalcReg(excelSheet.Cells(X, SN), excelSheet.Cells(X, IOCNO), excelSheet.Cells(X, CH))
                Else
                    excelSheet.Cells(X, Adr) = excelSheet.Cells(X, MbAdd)
                End If

                If Right(UCase(excelSheet.Cells(X, TNO)), 4) = Right(UCase(excelSheet.Cells(X + 1, TNO)), 4) And UCase(Right(excelSheet.Cells(X + 1, NODE), 4)) = "PORT" Then
                Else
                    c = c + 2
                End If
            Else

                If (UCase(excelSheet.Cells(X, ISA)) = "OPN" Or UCase(excelSheet.Cells(X, ISA)) = "CLS" Or UCase(excelSheet.Cells(X, ISA)) = "FOP" Or UCase(excelSheet.Cells(X, ISA)) = "FCL") And UCase(Right(excelSheet.Cells(X, NODE), 4)) = "STBD" Then
                    excelSheet.Cells(X, IasTN) = "RCV" & e & "_" & UCase(excelSheet.Cells(X, ISA))
                    If excelSheet.Cells(X, IOCNO) <> "" Then
                        excelSheet.Cells(X, Adr) = CalcReg(excelSheet.Cells(X, SN), excelSheet.Cells(X, IOCNO), excelSheet.Cells(X, CH))
                    Else
                        excelSheet.Cells(X, Adr) = excelSheet.Cells(X, MbAdd)
                    End If

                    If Right(UCase(excelSheet.Cells(X, TNO)), 4) = Right(UCase(excelSheet.Cells(X + 1, TNO)), 4) And UCase(Right(excelSheet.Cells(X + 1, NODE), 4)) = "STBD" Then
                    Else
                        e = e + 2
                    End If


                End If
            End If

            X = X + 1

        End While

    End Function

    Function tank()

        LoadIOLIst()

        Dim X = 9
        Dim a = 1
        Dim b = 1
        Dim c = 1
        Dim d = 1
        Dim e = 2
        While excelSheet.Cells(X, "A") <> ""

            ' find and asign tags for tank levels

            If Left(UCase(excelSheet.Cells(X, ISA)), 2) = "LI" And (Right(excelSheet.Cells(X, SN), 4) = "Port" Or Right(excelSheet.Cells(X, SN), 4) = "Stbd") Then
                excelSheet.Cells(X, IasTN) = "Tank" & Right(excelSheet.Cells(X, TNO), 3) & "_" & UCase(excelSheet.Cells(X, ISA))
                excelSheet.Cells(X, Adr) = CalcReg(excelSheet.Cells(X, SN), excelSheet.Cells(X, IOCNO), excelSheet.Cells(X, CH))


                If Right(UCase(excelSheet.Cells(X, TNO)), 3) = Right(UCase(excelSheet.Cells(X + 1, TNO)), 3) And (Right(excelSheet.Cells(X, SN), 4) = "Port" Or Right(excelSheet.Cells(X, SN), 4) = "Stbd") Then
                Else
                    d = d + 1
                End If
            End If



            X = X + 1
        End While


    End Function


    Function AlarmCsvGen()
        '
        ' Macro1 Macro
        ' Macro 10/27/2006 by Cristian Badescu
        '
        LoadIOLIst()

        Dim AlmNo As Integer
        Dim VectSP(0 To 82000) As String
        Dim VectSS(0 To 82000) As String


        Dim X = 9
        Dim a = 1
        Dim b = 1
        Dim c = 1
        Dim d = 1
        Dim i = 0
        Dim j = 0

        For i = 0 To 82000 ' cleen the vector's filds
            VectSP(i) = ""
            VectSS(j) = ""
        Next i
        X = 9
        i = 0
        j = 0

        While excelSheet.Cells(X, "A") <> ""

            If Left(excelSheet.Cells(X, IasTN), 3) = "Alm" Then

                AlmNo = Right(excelSheet.Cells(X, IasTN), 4)

                If AlmNo < MaxAlm Then
                    VectSP(i) = AlmNo
                    VectSP(i + 1) = SignalType(excelSheet.Cells(X, ST))
                    VectSP(i + 2) = excelSheet.Cells(X, AG)
                    VectSP(i + 3) = excelSheet.Cells(X, ALD)
                    VectSP(i + 4) = excelSheet.Cells(X, BG)
                    VectSP(i + 5) = excelSheet.Cells(X, AHH)
                    VectSP(i + 6) = excelSheet.Cells(X, AH)
                    VectSP(i + 7) = excelSheet.Cells(X, AL)

                    If VectSP(i + 1) = 1 Then
                        Select Case excelSheet.Cells(X, NVAL)
                            Case "NC"
                                VectSP(i + 8) = 0
                            Case "NO"
                                VectSP(i + 8) = 1
                        End Select
                    Else
                        VectSP(i + 8) = excelSheet.Cells(X, ALL)
                    End If

                    VectSP(i + 9) = excelSheet.Cells(X, Adr)
                    VectSP(i + 10) = excelSheet.Cells(X, RMin)
                    VectSP(i + 11) = excelSheet.Cells(X, RMax)
                    VectSP(i + 12) = excelSheet.Cells(X, Emin)
                    VectSP(i + 13) = excelSheet.Cells(X, EMax)
                    i = i + 14
                Else
                    VectSS(j) = AlmNo
                    VectSS(j + 1) = SignalType(excelSheet.Cells(X, ST))
                    VectSS(j + 2) = excelSheet.Cells(X, AG)
                    VectSS(j + 3) = excelSheet.Cells(X, ALD)
                    VectSS(j + 4) = excelSheet.Cells(X, BG)
                    VectSS(j + 5) = excelSheet.Cells(X, AHH)
                    VectSS(j + 6) = excelSheet.Cells(X, AH)
                    VectSS(j + 7) = excelSheet.Cells(X, AL)

                    If VectSS(j + 1) = 1 Then
                        Select Case excelSheet.Cells(X, NVAL)
                            Case "NC"
                                VectSS(j + 8) = 0
                            Case "NO"
                                VectSS(j + 8) = 1
                        End Select
                    Else
                        VectSS(j + 8) = excelSheet.Cells(X, ALL)
                    End If

                    VectSS(j + 9) = excelSheet.Cells(X, Adr)
                    VectSS(j + 10) = excelSheet.Cells(X, RMin)
                    VectSS(j + 11) = excelSheet.Cells(X, RMax)
                    VectSS(j + 12) = excelSheet.Cells(X, Emin)
                    VectSS(j + 13) = excelSheet.Cells(X, EMax)
                    j = j + 14
                End If

            End If
            X = X + 1
        End While

        Dim lContrX As Worksheet
        lContrX = Excel.ActiveWorkbook.Sheets("Contr01PS")


        X = 4
        i = 0
        While VectSP(i) <> ""
            'lContrX.Cells(X, "A") = VectSP(i)
            lContrX.Cells(X, "B") = Val(VectSP(i + 1))
            lContrX.Cells(X, "C") = Val(VectSP(i + 2))
            lContrX.Cells(X, "D") = Val(VectSP(i + 3))
            lContrX.Cells(X, "F") = Val(VectSP(i + 4))
            lContrX.Cells(X, "H") = Val(VectSP(i + 5))
            lContrX.Cells(X, "I") = Val(VectSP(i + 6))
            lContrX.Cells(X, "J") = Val(VectSP(i + 7))
            lContrX.Cells(X, "L") = Val(VectSP(i + 10))
            lContrX.Cells(X, "N") = Val(VectSP(i + 12))
            lContrX.Cells(X, "M") = Val(VectSP(i + 11))
            lContrX.Cells(X, "O") = Val(VectSP(i + 13))

            lContrX.Cells(X, "K") = Val(VectSP(i + 8))

            lContrX.Cells(X, "P") = VectSP(i + 9)

            X = X + 1
            i = i + 14
        End While

        lContrX = Excel.ActiveWorkbook.Sheets("Contr02SB")


        X = 4
        i = 0
        While VectSS(i) <> ""
            'lContrX.Cells(X, "A") = VectSS(i)
            lContrX.Cells(X, "B") = Val(VectSS(i + 1))
            lContrX.Cells(X, "C") = Val(VectSS(i + 2))
            lContrX.Cells(X, "D") = Val(VectSS(i + 3))
            lContrX.Cells(X, "F") = Val(VectSS(i + 4))
            lContrX.Cells(X, "H") = Val(VectSS(i + 5))
            lContrX.Cells(X, "I") = Val(VectSS(i + 6))
            lContrX.Cells(X, "J") = Val(VectSS(i + 7))
            lContrX.Cells(X, "L") = Val(VectSS(i + 10))
            lContrX.Cells(X, "N") = Val(VectSS(i + 12))
            lContrX.Cells(X, "M") = Val(VectSS(i + 11))
            lContrX.Cells(X, "O") = Val(VectSS(i + 13))

            lContrX.Cells(X, "K") = Val(VectSS(i + 8))

            lContrX.Cells(X, "P") = VectSS(i + 9)

            X = X + 1
            i = i + 14
        End While


    End Function

    Function SignalType(SgType As String) As Integer

        Select Case SgType

            Case "DI"
                SignalType = 1
            Case "DIC"
                SignalType = 1
            Case "AIC02"
                SignalType = 2
            Case "AIC"
                SignalType = 2
            Case "AI"
                SignalType = 2

        End Select

    End Function

    Function FillAker_dbf()

        LoadIOLIst()

        Dim AlmNo As Integer
        Dim VectSP(0 To 20000) As String
        Dim VectSS(0 To 20000) As String

        Dim X = 9
        Dim a = 1
        Dim b = 1
        Dim c = 1
        Dim d = 1
        Dim i = 0
        Dim j = 0

        For i = 0 To 20000 ' cleen the vector's filds
            VectSP(i) = ""
            VectSS(j) = ""
        Next i
        i = 1
        j = 1

        While excelSheet.Cells(X, "A") <> ""

            If Left(excelSheet.Cells(X, IasTN), 3) = "Alm" Then


                AlmNo = Right(excelSheet.Cells(X, IasTN), 4)

                If AlmNo < MaxAlm Then
                    VectSP(i) = AlmNo
                    VectSP(i + 1) = excelSheet.Cells(X, TN)
                    VectSP(i + 2) = excelSheet.Cells(X, DES)
                    If excelSheet.Cells(X, UN) <> "" Then
                        VectSP(i + 3) = excelSheet.Cells(X, UN)
                    Else
                        VectSP(i + 3) = "-"
                    End If

                    If SignalType(excelSheet.Cells(X, ST)) = 1 Then
                        Select Case excelSheet.Cells(X, NVAL)
                            Case "NC"
                                VectSP(i + 4) = "Normal"
                                VectSP(i + 5) = "Alarm"
                            Case "NO"
                                VectSP(i + 4) = "Alarm"
                                VectSP(i + 5) = "Normal"
                        End Select
                    End If

                    i = i + 6
                    VectSP(0) = Val(VectSP(0)) + 1
                Else
                    VectSS(j) = AlmNo
                    VectSS(j + 1) = excelSheet.Cells(X, TN)
                    VectSS(j + 2) = excelSheet.Cells(X, DES)
                    If excelSheet.Cells(X, UN) <> "" Then
                        VectSS(j + 3) = excelSheet.Cells(X, UN)
                    Else
                        VectSS(j + 3) = "-"
                    End If

                    If SignalType(excelSheet.Cells(X, ST)) = 1 Then
                        Select Case excelSheet.Cells(X, NVAL)
                            Case "NC"
                                VectSS(j + 4) = "Normal"
                                VectSS(j + 5) = "Alarm"
                            Case "NO"
                                VectSS(j + 4) = "Alarm"
                                VectSS(j + 5) = "Normal"
                        End Select
                    End If

                    j = j + 6

                    VectSS(0) = Val(VectSS(0)) + 1
                End If

            End If
            X = X + 1
        End While

        Dim lAKER_DBF_WB As Workbook
        lAKER_DBF_WB = Excel.Workbooks.Open(Aker_PATH)


        Dim lAker_sheet As Worksheet
        lAker_sheet = lAKER_DBF_WB.Sheets("AKER")


        X = 1
        While lAker_sheet.Cells(X, "A") <> "Alm0001_NAME"
            X = X + 1
        End While
        Dim FirstAlarm = X
        ' PS alarms

        i = 1
        a = 1
        X = FirstAlarm
        For i = 1 To 6000 * MaxAlm / 1000 Step 6


            If VectSP(i) <> "" Then
                lAker_sheet.Cells(X, "B") = VectSP(i + 1)
                lAker_sheet.Cells(X + 2000 * MaxAlm / 1000, "B") = VectSP(i + 2)
                lAker_sheet.Cells(X + 4000 * MaxAlm / 1000, "B") = VectSP(i + 3)
            Else
                lAker_sheet.Cells(X, "B") = "Alm" & Format(VectSP(0) + a, "000#") & "_NAME"
                lAker_sheet.Cells(X + 2000 * MaxAlm / 1000, "B") = "Alm" & Format(VectSP(0) + a, "000#") & "_DESC"
                lAker_sheet.Cells(X + 4000 * MaxAlm / 1000, "B") = "-"
                a = a + 1
            End If

            If VectSP(i + 4) <> "" Then
                lAker_sheet.Cells(X + 6000 * MaxAlm / 1000, "B") = VectSP(i + 4)
                lAker_sheet.Cells(X + 8000 * MaxAlm / 1000, "B") = VectSP(i + 5)
            Else
                lAker_sheet.Cells(X + 6000 * MaxAlm / 1000, "B") = "Alarm"
                lAker_sheet.Cells(X + 8000 * MaxAlm / 1000, "B") = "Normal"
            End If

            X = X + 1

        Next i

        'SB Alarms
        i = 1
        a = 1
        X = FirstAlarm

        For i = 1 To 4000 * MaxAlm / 1000 Step 6
            If VectSS(i) <> "" Then
                lAker_sheet.Cells(X + 1000 * MaxAlm / 1000, "B") = VectSS(i + 1)
                lAker_sheet.Cells(X + 3000 * MaxAlm / 1000, "B") = VectSS(i + 2)
                lAker_sheet.Cells(X + 5000 * MaxAlm / 1000, "B") = VectSS(i + 3)
            Else
                lAker_sheet.Cells(X + 1000 * MaxAlm / 1000, "B") = "Alm" & VectSS(0) + a + MaxAlm & "_NAME"
                lAker_sheet.Cells(X + 3000 * MaxAlm / 1000, "B") = "Alm" & VectSS(0) + a + MaxAlm & "_DESC"
                lAker_sheet.Cells(X + 5000 * MaxAlm / 1000, "B") = "-"
                a = a + 1
            End If

            If VectSS(i + 4) <> "" Then
                excelSheet.Cells(X + 7000 * MaxAlm / 1000, "B") = VectSS(i + 4)
                excelSheet.Cells(X + 9000 * MaxAlm / 1000, "B") = VectSS(i + 5)
            Else
                excelSheet.Cells(X + 7000 * MaxAlm / 1000, "B") = "Alarm"
                excelSheet.Cells(X + 9000 * MaxAlm / 1000, "B") = "Normal"
            End If

            X = X + 1

        Next i


    End Function

    Function IsaGraf()
        '
        ' Macro1 Macro
        ' Macro 10/27/2006 by Cristian Badescu
        '

        LoadIOLIst()
        Dim VectIsaG(0 To 50000) As String

        Dim X = 9
        Dim a = 1
        Dim b = 1
        Dim c = 1
        Dim d = 1
        Dim i = 0
        Dim j = 0

        For i = 0 To 50000 ' cleen the vector's filds
            VectIsaG(i) = ""
        Next i
        i = 0
        j = 0

        While excelSheet.Cells(X, "A") <> ""

            If excelSheet.Cells(X, "AW") <> "" Then

                If excelSheet.Cells(X, "AW") = "*" Then
                    VectIsaG(i) = excelSheet.Cells(X, "D")
                Else
                    VectIsaG(i) = excelSheet.Cells(X, "AW")
                End If

                If UCase(Right(excelSheet.Cells(X, "AC"), 4)) = "PORT" Then
                    VectIsaG(i + 1) = " Contr01Port"
                    VectIsaG(i + 2) = 1
                Else
                    If UCase(Right(excelSheet.Cells(X, "AC"), 4)) = "STBD" Then
                        VectIsaG(i + 1) = " Contr02Stbd"
                        VectIsaG(i + 2) = 2
                    End If

                End If
                Select Case excelSheet.Cells(X, "I")

                    Case "DI"
                        VectIsaG(i + 3) = 10
                    Case "DO"
                        VectIsaG(i + 3) = 11
                    Case "AI"
                        VectIsaG(i + 3) = 0
                    Case "AO"
                        VectIsaG(i + 3) = 1
                End Select
                '               VectIsaG(I + 3) = SignalType(excelSheet.Cells(x, "I"))
                VectIsaG(i + 4) = Right(excelSheet.Cells(X, "AY"), 4)
                If Left(excelSheet.Cells(X, "AY"), 1) = "A" Then
                    VectIsaG(i + 5) = " short"
                Else
                    VectIsaG(i + 5) = " discrete"
                End If

                i = i + 6
            End If

            X = X + 1
        End While


        Dim lISagraf As Worksheet
        lISagraf = Excel.ActiveWorkbook.Sheets("IsaGraf")

        X = 2
        i = 0
        While VectIsaG(i) <> ""
            lISagraf.Cells(X, "A") = VectIsaG(i)
            lISagraf.Cells(X, "B") = VectIsaG(i + 1)
            lISagraf.Cells(X, "C") = VectIsaG(i + 2)
            lISagraf.Cells(X, "D") = VectIsaG(i + 3)
            lISagraf.Cells(X, "E") = VectIsaG(i + 4)
            lISagraf.Cells(X, "F") = VectIsaG(i + 5)

            X = X + 1
            i = i + 6
        End While



    End Function
    Function AlarmAddress()
        '
        ' Macro1 Macro
        ' Macro 10/27/2006 by Cristian Badescu

        LoadIOLIst()
        Dim Vect(0 To 40000) As String

        For i = 0 To 40000 ' cleen the vector's filds
            Vect(i) = ""
        Next i

        Dim X = 9
        Dim a = 0


        While excelSheet.Cells(X, "A") <> ""

            Vect(a) = excelSheet.Cells(X, "D") & "_" & excelSheet.Cells(X, "E")
            Vect(a + 1) = excelSheet.Cells(X, RMin)  'A Range Min
            Vect(a + 2) = excelSheet.Cells(X, RMax)  'A Range Max
            Vect(a + 3) = excelSheet.Cells(X, UN)  'A Unit
            Vect(a + 4) = excelSheet.Cells(X, ALD)  'A Delay
            Vect(a + 5) = excelSheet.Cells(X, ALL)  'ALL
            Vect(a + 6) = excelSheet.Cells(X, AL)  'AL
            Vect(a + 7) = excelSheet.Cells(X, AH)  'AH
            Vect(a + 8) = excelSheet.Cells(X, AHH)  'AHH
            Vect(a + 9) = excelSheet.Cells(X, Adr) 'Address
            Vect(a + 10) = excelSheet.Cells(X, ST) 'Sensor Type


            a = a + 11
            X = X + 1
        End While

        Dim lAlarms As Worksheet
        lAlarms = Excel.ActiveWorkbook.Sheets("Alarms")


        X = 2
        a = 0

        While excelSheet.Cells(X, "A") <> ""

            For i = 0 To 40000 Step 11 ' cleen the vector's filds

                If excelSheet.Cells(X, "A") = Vect(i) Then

                    If Vect(i + 1) <> "" Then
                        excelSheet.Cells(X, "Q") = Vect(i + 1)
                    End If
                    If Vect(i + 2) <> "" Then
                        excelSheet.Cells(X, "R") = Vect(i + 2)
                    End If
                    If Vect(i + 3) <> "" Then
                        excelSheet.Cells(X, "C") = Vect(i + 3)
                    End If
                    If Vect(i + 4) <> "" Then
                        excelSheet.Cells(X, "G") = Vect(i + 4)
                    End If
                    If Vect(i + 8) <> "" Then
                        excelSheet.Cells(X, "K") = Vect(i + 8)
                    End If
                    If Vect(i + 7) <> "" Then
                        excelSheet.Cells(X, "L") = Vect(i + 7)
                    End If
                    If Vect(i + 6) <> "" Then
                        excelSheet.Cells(X, "M") = Vect(i + 6)
                    End If
                    If Vect(i + 5) <> "" Then
                        excelSheet.Cells(X, "N") = Vect(i + 5)
                    End If
                    If Vect(i + 9) <> "" Then
                        excelSheet.Cells(X, "T") = Vect(i + 9)
                    End If

                    Select Case UCase(Trim(Vect(i + 10)))

                        Case "4-20MA"

                            excelSheet.Cells(X, "O") = 0
                            excelSheet.Cells(X, "P") = 32767

                        Case "PT-100"

                            If Vect(i + 1) <> "" Then
                                excelSheet.Cells(X, "O") = Vect(i + 1)
                            End If

                            If Vect(i + 2) <> "" Then
                                excelSheet.Cells(X, "P") = Vect(i + 2)
                            End If

                        Case "MODBUS"

                            If Vect(i + 1) <> "" Then
                                excelSheet.Cells(X, "O") = Vect(i + 1)
                            End If
                            If Vect(i + 2) <> "" Then
                                excelSheet.Cells(X, "P") = Vect(i + 2)
                            End If

                        Case "NMEA"

                            If Vect(i + 1) <> "" Then
                                excelSheet.Cells(X, "O") = Vect(i + 1)
                            End If
                            If Vect(i + 2) <> "" Then
                                excelSheet.Cells(X, "P") = Vect(i + 2)
                            End If

                        Case Else
                            excelSheet.Cells(X, "O") = 0
                            excelSheet.Cells(X, "P") = 100

                    End Select
                    GoTo line1
                End If
            Next i
line1:
            X = X + 1
        End While

    End Function

End Module
