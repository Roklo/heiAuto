Module modCustom


    'THIS IS CUSTOM FUNCTIONS that you should modify to suit your specific project
    Public Function CustomProjectName() As String
        CustomProjectName = "NB740" ' This is the citect project name
    End Function


    Public Function CustomAdjustSFINumber(piGenieName As String, piInputSFI As String) As String
        Dim lNewVal As String

        If Not InStr(piInputSFI, "_") Then
            Select Case piGenieName
                Case "lal"
                    'Convert from "722L101" to "722_L101_LAL"
                    lNewVal = Left(piGenieName, 3) + "_" + Right(piGenieName, 4) + "_LAL"
                Case "lah"
                    lNewVal = Left(piGenieName, 3) + "_" + Right(piGenieName, 4) + "_LAH"
                Case "rcv_90"
                    lNewVal = Left(piGenieName, 3) + "_V0" + Right(piGenieName, 3) + "_OPN"
                Case "alarm_light_square"
                    lNewVal = Left(piGenieName, 3) + "_L" + Right(piGenieName, 3) + "_LAH"
            End Select
        End If

        CustomAdjustSFINumber = lNewVal
    End Function



End Module
