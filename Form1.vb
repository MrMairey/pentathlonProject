Public Class Form1
    Dim currEvent As Integer '0 = 200m, 1 = 1500m, 2 = Discus, 3 = Javelin, 4 = Long Jump
    Dim results200m(2) As Decimal
    Dim results1500m(2) As Decimal
    Dim resultsJavelin(2) As Decimal
    Dim resultsLJump(2) As Decimal
    Dim resultsDiscus(2) As Decimal
    Dim strFileName As String
    Dim userPath As String
    Dim comp1Points(4) As Decimal
    Dim comp2Points(4) As Decimal
    Dim comp3Points(4) As Decimal
    Dim overviewPointsCache As Decimal
    Dim success As Boolean = True
    Dim comp1ChkFlag(4) As Boolean
    Dim comp2ChkFlag(4) As Boolean
    Dim comp3ChkFlag(4) As Boolean
    Dim comp1Score(4) As Decimal
    Dim comp2Score(4) As Decimal
    Dim comp3Score(4) As Decimal
    Dim resultsReOrderCache(2) As Decimal
    Dim antisoftlock As Integer
    Dim pointsTotal(2) As Decimal
    Dim negNumTest As Decimal
    Dim junkData As String
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For i As Integer = 0 To 4
            comp1ChkFlag(i) = True
            comp2ChkFlag(i) = True
            comp3ChkFlag(i) = True 'sets the flags for the check boxes to true by default
        Next
        userPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) 'current user's desktop path
        strFileName = IO.Path.Combine(userPath, "PentathlonResults.txt") 'current user's desktop path + the "PentathlonResults.txt"
    End Sub
    Public Sub overviewTab()
        If comp1Score(currEvent) = 0 Then 'if 0 is entered for the score, DNF is activated by setting the current points for the current event to 0
            comp1Points(currEvent) = 0
        End If
        If comp2Score(currEvent) = 0 Then
            comp2Points(currEvent) = 0
        End If
        If comp3Score(currEvent) = 0 Then
            comp3Points(currEvent) = 0
        End If
        'comp1
        overviewPointsCache = overviewPoints1Lbl.Text
        overviewPointsCache = overviewPointsCache + comp1Points(currEvent)
        pointsTotal(0) = overviewPointsCache
        overviewPoints1Lbl.Text = overviewPointsCache 'Updates the current total points counter as well as updating the array with the same value
        'comp2
        overviewPointsCache = overviewPoints2Lbl.Text
        overviewPointsCache = overviewPointsCache + comp2Points(currEvent)
        pointsTotal(1) = overviewPointsCache
        overviewPoints2Lbl.Text = overviewPointsCache
        'comp3
        overviewPointsCache = overviewPoints3Lbl.Text
        overviewPointsCache = overviewPointsCache + comp3Points(currEvent)
        pointsTotal(2) = overviewPointsCache
        overviewPoints3Lbl.Text = overviewPointsCache
        If pointsTotal(0) = pointsTotal(1) And pointsTotal(0) = pointsTotal(2) Then 'decisions on what placement each competitor should get -vvvvvv-
            Comp1PlcLbl.BackColor = Color.Gold
            Comp1PlcLbl.Text = "1st"
            Comp2PlcLbl.BackColor = Color.Gold 'if all points are equal ~
            Comp2PlcLbl.Text = "1st"
            Comp3PlcLbl.BackColor = Color.Gold
            Comp3PlcLbl.Text = "1st"
        ElseIf pointsTotal(0) = pointsTotal(1) And pointsTotal(0) > pointsTotal(2) Then ' comp1 and comp2 tie for first
            Comp1PlcLbl.BackColor = Color.Gold
            Comp1PlcLbl.Text = "1st"
            Comp2PlcLbl.BackColor = Color.Gold
            Comp2PlcLbl.Text = "1st"
        ElseIf pointsTotal(0) = pointsTotal(2) And pointsTotal(0) > pointsTotal(1) Then ' comp1 and comp3 tie for first
            Comp1PlcLbl.BackColor = Color.Gold
            Comp1PlcLbl.Text = "1st"
            Comp3PlcLbl.BackColor = Color.Gold
            Comp3PlcLbl.Text = "1st"
        ElseIf pointsTotal(1) = pointsTotal(2) And pointsTotal(1) > pointsTotal(0) Then ' comp2 and comp3 tie for first
            Comp2PlcLbl.BackColor = Color.Gold
            Comp2PlcLbl.Text = "1st"
            Comp3PlcLbl.BackColor = Color.Gold
            Comp3PlcLbl.Text = "1st"
        End If
        If pointsTotal(0) = pointsTotal(1) And pointsTotal(0) < pointsTotal(2) Then ' comp1 and comp2 tie for second
            Comp1PlcLbl.BackColor = Color.Silver
            Comp1PlcLbl.Text = "2nd"
            Comp2PlcLbl.BackColor = Color.Silver
            Comp2PlcLbl.Text = "2nd"
        ElseIf pointsTotal(0) = pointsTotal(2) And pointsTotal(0) < pointsTotal(1) Then ' comp1 and comp3 tie for second
            Comp1PlcLbl.BackColor = Color.Silver
            Comp1PlcLbl.Text = "2nd"
            Comp3PlcLbl.BackColor = Color.Silver
            Comp3PlcLbl.Text = "2nd"
        ElseIf pointsTotal(1) = pointsTotal(2) And pointsTotal(1) < pointsTotal(0) Then ' comp2 and comp3 tie for second
            Comp2PlcLbl.BackColor = Color.Silver
            Comp2PlcLbl.Text = "2nd"
            Comp3PlcLbl.BackColor = Color.Silver
            Comp3PlcLbl.Text = "2nd"
        End If
        If pointsTotal(0) > pointsTotal(1) And pointsTotal(0) > pointsTotal(2) Then 'only comp1 gets first
            Comp1PlcLbl.BackColor = Color.Gold
            Comp1PlcLbl.Text = "1st"
        ElseIf pointsTotal(1) > pointsTotal(0) And pointsTotal(1) > pointsTotal(2) Then 'only comp2 gets first
            Comp2PlcLbl.BackColor = Color.Gold
            Comp2PlcLbl.Text = "1st"
        ElseIf pointsTotal(2) > pointsTotal(0) And pointsTotal(2) > pointsTotal(1) Then 'only comp3 gets first
            Comp3PlcLbl.BackColor = Color.Gold
            Comp3PlcLbl.Text = "1st"
        End If
        If pointsTotal(0) >= pointsTotal(1) Xor pointsTotal(0) >= pointsTotal(2) Then 'only comp1 gets second
            Comp1PlcLbl.BackColor = Color.Silver
            Comp1PlcLbl.Text = "2nd"
        ElseIf pointsTotal(1) >= pointsTotal(0) Xor pointsTotal(1) >= pointsTotal(2) Then 'only comp2 gets second
            Comp2PlcLbl.BackColor = Color.Silver
            Comp2PlcLbl.Text = "2nd"
        ElseIf pointsTotal(2) >= pointsTotal(0) Xor pointsTotal(2) >= pointsTotal(1) Then 'only comp3 gets second
            Comp3PlcLbl.BackColor = Color.Silver
            Comp3PlcLbl.Text = "2nd"
        End If
        If pointsTotal(0) < pointsTotal(1) And pointsTotal(0) < pointsTotal(2) Then 'comp1 gets third
            Comp1PlcLbl.BackColor = Color.Bisque
            Comp1PlcLbl.Text = "3rd"
        ElseIf pointsTotal(1) < pointsTotal(0) And pointsTotal(1) < pointsTotal(2) Then 'comp2 gets third
            Comp2PlcLbl.BackColor = Color.Bisque
            Comp2PlcLbl.Text = "3rd"
        ElseIf pointsTotal(2) < pointsTotal(0) And pointsTotal(2) < pointsTotal(1) Then 'comp3 gets third
            Comp3PlcLbl.BackColor = Color.Bisque
            Comp3PlcLbl.Text = "3rd"
        End If
    End Sub
    Private Sub confPartBtn_Click(sender As Object, e As EventArgs) Handles confPartBtn.Click
        compName1txt.Enabled = False 'disable name text boxes
        compName2txt.Enabled = False
        compName3txt.Enabled = False
        compName1txt.Text = compName1txt.Text.ToUpperInvariant 'set names to uppercase
        compName2txt.Text = compName2txt.Text.ToUpperInvariant
        compName3txt.Text = compName3txt.Text.ToUpperInvariant
        overviewName1Lbl.Text = compName1txt.Text 'set overview section names
        overviewName2Lbl.Text = compName2txt.Text
        overviewName3Lbl.Text = compName3txt.Text
        confPartBtn.Text = "Participants Confirmed" 'change button text
        eventTabControl.Enabled = True 'enable tab control, allowing results to be entered
        confPartBtn.Enabled = False 'disable name confirmation button
        instructLbl.Text = "Please enter values into each text box for a chosen event and press 'Validate and save results' when done. Check the 'DNF' check-box or type '0' as the result if the competitor does not finish or compete"
    End Sub 'change instruction label^^^
    '200m Checks
    Private Sub Conf200mBtn_Click(sender As Object, e As EventArgs) Handles Conf200mBtn.Click
        currEvent = 0 'tells program what event is currently taking place
        Try 'trap any errors
            If comp1ChkFlag(currEvent) = True Then ' if the DNF checkbox isnt checked:
                negNumTest = comp1200mtxt.Text
                If negNumTest < 0 Then 'if the entered value is a negative number:
                    success = "" 'induce an error
                End If
                results200m(0) = comp1200mtxt.Text 'place the entered results into the array to later be sorted
                comp1Score(currEvent) = comp1200mtxt.Text 'place the entered results into another array unique to the competitor
            Else comp1200mtxt.Text = "0" ' if the DNF checkbox is checked, the competitor's result is set to 0
            End If
            If comp2ChkFlag(currEvent) = True Then
                negNumTest = comp2200mtxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                results200m(1) = comp2200mtxt.Text
                comp2Score(currEvent) = comp2200mtxt.Text
            Else comp2200mtxt.Text = "0"
            End If
            If comp3ChkFlag(currEvent) = True Then
                negNumTest = comp3200mtxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                results200m(2) = comp3200mtxt.Text
                comp3Score(currEvent) = comp3200mtxt.Text
            Else comp3200mtxt.Text = "0"
            End If
            Conf200mBtn.Enabled = False 'disable the event confirmation button
        Catch ex As Exception 'catch any errors and return line 178 & 179
            MsgBox("Please input a valid result with the format:'[seconds].[milliseconds]'", , "Error") 'error message
            success = False 'prevent the code from going any further if an error occurs
        End Try
        If success = False Then
            success = True 'changes the success boolean to true then ends the code, allowing for infinite attempts. this is only ever false if an error occurs during that single run.
        Else
            Array.Sort(results200m) 'sort results array in ascending order
            antisoftlock = 0 'essentially allowing a for statement and a doWhile. antisoftlock prevents an infinite loop in the 'for' statement if all the inputs for that event are DNF or '0'
            Do While results200m(0) = 0 And antisoftlock < 2 'of there is a value of '0' in the array,  the 0 value moves to the last place of the array vvvvvv
                resultsReOrderCache(0) = results200m(0)
                resultsReOrderCache(1) = results200m(1)
                resultsReOrderCache(2) = results200m(2)
                results200m(2) = resultsReOrderCache(0)
                results200m(1) = resultsReOrderCache(2)
                results200m(0) = resultsReOrderCache(1)
                antisoftlock = antisoftlock + 1 'counts how many times the do while has occured
            Loop
            Select Case comp1Score(currEvent) 'if competitor1's score matches the:
                Case results200m(0) ' first value of the array
                    comp1Points(currEvent) = 3.1 'is set as 3.1 due to the need for the first place player aggregation 
                Case results200m(1) 'second value of the array
                    comp1Points(currEvent) = 2
                Case results200m(2) ' the third value of the array
                    comp1Points(currEvent) = 1
            End Select
            Select Case comp2Score(currEvent)
                Case results200m(0)
                    comp2Points(currEvent) = 3.1
                Case results200m(1)
                    comp2Points(currEvent) = 2
                Case results200m(2)
                    comp2Points(currEvent) = 1
            End Select
            Select Case comp3Score(currEvent)
                Case results200m(0)
                    comp3Points(currEvent) = 3.1
                Case results200m(1)
                    comp3Points(currEvent) = 2
                Case results200m(2)
                    comp3Points(currEvent) = 1
            End Select
            comp1200mtxt.Enabled = False 'disables all screen elements of the current event
            comp2200mtxt.Enabled = False
            comp3200mtxt.Enabled = False
            comp1200mChk.Enabled = False
            comp2200mChk.Enabled = False
            comp3200mChk.Enabled = False
            If ProgressBar1.Value <= 80 Then 'tick up the progress bar cause its cool
                ProgressBar1.Value = ProgressBar1.Value + 20
            End If
            overviewTab() 'open overview tab sub-program
        End If
    End Sub
    Private Sub Comp1200mChk_CheckedChanged(sender As Object, e As EventArgs) Handles comp1200mChk.CheckedChanged
        If comp1ChkFlag(0) = False Then 'toggles the associated boolean along with the 'Enabled' property of the associated text box
            comp1200mtxt.Enabled = True
            comp1ChkFlag(0) = True
        Else
            comp1200mtxt.Enabled = False
            comp1ChkFlag(0) = False
        End If
    End Sub
    Private Sub Comp2200mChk_CheckedChanged(sender As Object, e As EventArgs) Handles comp2200mChk.CheckedChanged
        If comp2ChkFlag(0) = False Then
            comp2200mtxt.Enabled = True
            comp2ChkFlag(0) = True
        Else
            comp2200mtxt.Enabled = False
            comp2ChkFlag(0) = False
        End If
    End Sub
    Private Sub Comp3200mChk_CheckedChanged(sender As Object, e As EventArgs) Handles comp3200mChk.CheckedChanged
        If comp3ChkFlag(0) = False Then
            comp3200mtxt.Enabled = True
            comp3ChkFlag(0) = True
        Else
            comp3200mtxt.Enabled = False
            comp3ChkFlag(0) = False
        End If
    End Sub
    '1500m Checks
    Private Sub Conf1500mBtn_Click(sender As Object, e As EventArgs) Handles Conf1500mBtn.Click
        currEvent = 1
        Try
            If comp1ChkFlag(currEvent) = True Then
                negNumTest = comp11500mtxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                results1500m(0) = comp11500mtxt.Text
                comp1Score(currEvent) = comp11500mtxt.Text
            Else comp11500mtxt.Text = "0"
            End If
            If comp2ChkFlag(currEvent) = True Then
                negNumTest = comp21500mtxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                results1500m(1) = comp21500mtxt.Text
                comp2Score(currEvent) = comp21500mtxt.Text
            Else comp21500mtxt.Text = "0"
            End If
            If comp3ChkFlag(currEvent) = True Then
                negNumTest = comp31500mtxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                results1500m(2) = comp31500mtxt.Text
                comp3Score(currEvent) = comp31500mtxt.Text
            Else comp31500mtxt.Text = "0"
            End If
            Conf1500mBtn.Enabled = False
        Catch ex As Exception
            MsgBox("Please input a valid result with the format:'[seconds].[milliseconds]'", , "Error")
            success = False
        End Try
        If success = False Then
            success = True
        Else
            Array.Sort(results1500m)
            antisoftlock = 0
            Do While results1500m(0) = 0 And antisoftlock < 2
                resultsReOrderCache(0) = results1500m(0)
                resultsReOrderCache(1) = results1500m(1)
                resultsReOrderCache(2) = results1500m(2)
                results1500m(2) = resultsReOrderCache(0)
                results1500m(1) = resultsReOrderCache(2)
                results1500m(0) = resultsReOrderCache(1)
                antisoftlock = antisoftlock + 1
            Loop
            Select Case comp1Score(currEvent)
                Case results1500m(0)
                    comp1Points(currEvent) = 3
                Case results1500m(1)
                    comp1Points(currEvent) = 2
                Case results1500m(2)
                    comp1Points(currEvent) = 1
            End Select
            Select Case comp2Score(currEvent)
                Case results1500m(0)
                    comp2Points(currEvent) = 3
                Case results1500m(1)
                    comp2Points(currEvent) = 2
                Case results1500m(2)
                    comp2Points(currEvent) = 1
            End Select
            Select Case comp3Score(currEvent)
                Case results1500m(0)
                    comp3Points(currEvent) = 3
                Case results1500m(1)
                    comp3Points(currEvent) = 2
                Case results1500m(2)
                    comp3Points(currEvent) = 1
            End Select
            comp11500mtxt.Enabled = False
            comp21500mtxt.Enabled = False
            comp31500mtxt.Enabled = False
            comp11500mChk.Enabled = False
            comp21500mChk.Enabled = False
            comp31500mChk.Enabled = False
            If ProgressBar1.Value <= 80 Then
                ProgressBar1.Value = ProgressBar1.Value + 20
            End If
            overviewTab()
        End If
    End Sub
    Private Sub Comp11500mChk_CheckedChanged(sender As Object, e As EventArgs) Handles comp11500mChk.CheckedChanged
        If comp1ChkFlag(1) = False Then
            comp11500mtxt.Enabled = True
            comp1ChkFlag(1) = True
        Else
            comp11500mtxt.Enabled = False
            comp1ChkFlag(1) = False
        End If
    End Sub
    Private Sub Comp21500mChk_CheckedChanged(sender As Object, e As EventArgs) Handles comp21500mChk.CheckedChanged
        If comp2ChkFlag(1) = False Then
            comp21500mtxt.Enabled = True
            comp2ChkFlag(1) = True
        Else
            comp21500mtxt.Enabled = False
            comp2ChkFlag(1) = False
        End If
    End Sub
    Private Sub Comp31500mChk_CheckedChanged(sender As Object, e As EventArgs) Handles comp31500mChk.CheckedChanged
        If comp3ChkFlag(1) = False Then
            comp31500mtxt.Enabled = True
            comp3ChkFlag(1) = True
        Else
            comp31500mtxt.Enabled = False
            comp3ChkFlag(1) = False
        End If
    End Sub
    'Discus Checks
    Private Sub confDiscusBtn_Click(sender As Object, e As EventArgs) Handles confDiscusBtn.Click
        currEvent = 2
        Try
            If comp1ChkFlag(currEvent) = True Then
                negNumTest = comp1Discustxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                resultsDiscus(0) = comp1Discustxt.Text
                comp1Score(currEvent) = comp1Discustxt.Text
            Else comp1Discustxt.Text = "0"
            End If
            If comp2ChkFlag(currEvent) = True Then
                negNumTest = comp2Discustxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                resultsDiscus(1) = comp2Discustxt.Text
                comp2Score(currEvent) = comp2Discustxt.Text
            Else comp2Discustxt.Text = "0"
            End If
            If comp3ChkFlag(currEvent) = True Then
                negNumTest = comp3Discustxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                resultsDiscus(2) = comp3Discustxt.Text
                comp3Score(currEvent) = comp3Discustxt.Text
            Else comp3Discustxt.Text = "0"
            End If
            confDiscusBtn.Enabled = False
        Catch ex As Exception
            MsgBox("Please input a valid result with either the format: '[centimeters].[millimeters]' or '[meters].[centimeters]'", , "Error")
            success = False
        End Try
        If success = False Then
            success = True
        Else
            Array.Sort(resultsDiscus)
            Select Case comp1Score(currEvent)
                Case resultsDiscus(0)
                    comp1Points(currEvent) = 1
                Case resultsDiscus(1)
                    comp1Points(currEvent) = 2
                Case resultsDiscus(2)
                    comp1Points(currEvent) = 3
            End Select
            Select Case comp2Score(currEvent)
                Case resultsDiscus(0)
                    comp2Points(currEvent) = 1
                Case resultsDiscus(1)
                    comp2Points(currEvent) = 2
                Case resultsDiscus(2)
                    comp2Points(currEvent) = 3
            End Select
            Select Case comp3Score(currEvent)
                Case resultsDiscus(0)
                    comp3Points(currEvent) = 1
                Case resultsDiscus(1)
                    comp3Points(currEvent) = 2
                Case resultsDiscus(2)
                    comp3Points(currEvent) = 3
            End Select
            comp1Discustxt.Enabled = False
            comp2Discustxt.Enabled = False
            comp3Discustxt.Enabled = False
            comp1DiscusChk.Enabled = False
            comp2DiscusChk.Enabled = False
            comp3DiscusChk.Enabled = False
            If ProgressBar1.Value <= 80 Then
                ProgressBar1.Value = ProgressBar1.Value + 20
            End If
            overviewTab()
        End If
    End Sub
    Private Sub Comp1DiscusChk_CheckedChanged(sender As Object, e As EventArgs) Handles comp1DiscusChk.CheckedChanged
        If comp1ChkFlag(2) = False Then
            comp1Discustxt.Enabled = True
            comp1ChkFlag(2) = True
        Else
            comp1Discustxt.Enabled = False
            comp1ChkFlag(2) = False
        End If
    End Sub
    Private Sub Comp2DiscusChk_CheckedChanged(sender As Object, e As EventArgs) Handles comp2DiscusChk.CheckedChanged
        If comp2ChkFlag(2) = False Then
            comp2Discustxt.Enabled = True
            comp2ChkFlag(2) = True
        Else
            comp2Discustxt.Enabled = False
            comp2ChkFlag(2) = False
        End If
    End Sub
    Private Sub Comp3DiscusChk_CheckedChanged(sender As Object, e As EventArgs) Handles comp3DiscusChk.CheckedChanged
        If comp3ChkFlag(2) = False Then
            comp3Discustxt.Enabled = True
            comp3ChkFlag(2) = True
        Else
            comp3Discustxt.Enabled = False
            comp3ChkFlag(2) = False
        End If
    End Sub
    'Jav Checks
    Private Sub confJavelinBtn_Click(sender As Object, e As EventArgs) Handles confJavelinBtn.Click
        currEvent = 3
        Try
            If comp1ChkFlag(currEvent) = True Then
                negNumTest = comp1Javelintxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                resultsJavelin(0) = comp1Javelintxt.Text
                comp1Score(currEvent) = comp1Javelintxt.Text
            Else comp1Javelintxt.Text = "0"
            End If
            If comp2ChkFlag(currEvent) = True Then
                negNumTest = comp2Javelintxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                resultsJavelin(1) = comp2Javelintxt.Text
                comp2Score(currEvent) = comp2Javelintxt.Text
            Else comp2Javelintxt.Text = "0"
            End If
            If comp3ChkFlag(currEvent) = True Then
                negNumTest = comp3Javelintxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                resultsJavelin(2) = comp3Javelintxt.Text
                comp3Score(currEvent) = comp3Javelintxt.Text
            Else comp3Javelintxt.Text = "0"
            End If
            confJavelinBtn.Enabled = False
        Catch ex As Exception
            MsgBox("Please input a valid result with either the format: '[centimeters].[millimeters]' or '[meters].[centimeters]'", , "Error")
            success = False
        End Try
        If success = False Then
            success = True
        Else
            Array.Sort(resultsJavelin)
            Select Case comp1Score(currEvent)
                Case resultsJavelin(0)
                    comp1Points(currEvent) = 1
                Case resultsJavelin(1)
                    comp1Points(currEvent) = 2
                Case resultsJavelin(2)
                    comp1Points(currEvent) = 3
            End Select
            Select Case comp2Score(currEvent)
                Case resultsJavelin(0)
                    comp2Points(currEvent) = 1
                Case resultsJavelin(1)
                    comp2Points(currEvent) = 2
                Case resultsJavelin(2)
                    comp2Points(currEvent) = 3
            End Select
            Select Case comp3Score(currEvent)
                Case resultsJavelin(0)
                    comp3Points(currEvent) = 1
                Case resultsJavelin(1)
                    comp3Points(currEvent) = 2
                Case resultsJavelin(2)
                    comp3Points(currEvent) = 3
            End Select
            comp1Javelintxt.Enabled = False
            comp2Javelintxt.Enabled = False
            comp3Javelintxt.Enabled = False
            Comp1JavelinChk.Enabled = False
            Comp2JavelinChk.Enabled = False
            Comp3JavelinChk.Enabled = False
            If ProgressBar1.Value <= 80 Then
                ProgressBar1.Value = ProgressBar1.Value + 20
            End If
            overviewTab()
        End If
    End Sub
    Private Sub Comp1JavelinChk_CheckedChanged(sender As Object, e As EventArgs) Handles Comp1JavelinChk.CheckedChanged
        If comp1ChkFlag(3) = False Then
            comp1Javelintxt.Enabled = True
            comp1ChkFlag(3) = True
        Else
            comp1Javelintxt.Enabled = False
            comp1ChkFlag(3) = False
        End If
    End Sub
    Private Sub Comp2JavelinChk_CheckedChanged(sender As Object, e As EventArgs) Handles Comp2JavelinChk.CheckedChanged
        If comp2ChkFlag(3) = False Then
            comp2Javelintxt.Enabled = True
            comp2ChkFlag(3) = True
        Else
            comp2Javelintxt.Enabled = False
            comp2ChkFlag(3) = False
        End If
    End Sub
    Private Sub Comp3JavelinChk_CheckedChanged(sender As Object, e As EventArgs) Handles Comp3JavelinChk.CheckedChanged
        If comp3ChkFlag(3) = False Then
            comp3Javelintxt.Enabled = True
            comp3ChkFlag(3) = True
        Else
            comp3Javelintxt.Enabled = False
            comp3ChkFlag(3) = False
        End If
    End Sub
    'LongJump Checks
    Private Sub confLongJumpBtn_Click(sender As Object, e As EventArgs) Handles confLJumpBtn.Click
        currEvent = 4
        Try
            If comp1ChkFlag(currEvent) = True Then
                negNumTest = comp1LJumptxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                resultsLJump(0) = comp1LJumptxt.Text
                comp1Score(currEvent) = comp1LJumptxt.Text
            Else comp1LJumptxt.Text = "0"
            End If
            If comp2ChkFlag(currEvent) = True Then
                negNumTest = comp2LJumptxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                resultsLJump(1) = comp2LJumptxt.Text
                comp2Score(currEvent) = comp2LJumptxt.Text
            Else comp2LJumptxt.Text = "0"
            End If
            If comp3ChkFlag(currEvent) = True Then
                negNumTest = comp3LJumptxt.Text
                If negNumTest < 0 Then
                    success = ""
                End If
                resultsLJump(2) = comp3LJumptxt.Text
                comp3Score(currEvent) = comp3LJumptxt.Text
            Else comp3LJumptxt.Text = "0"
            End If
            confLJumpBtn.Enabled = False
        Catch ex As Exception
            MsgBox("Please input a valid result with either the format: '[centimeters].[millimeters]' or '[meters].[centimeters]'", , "Error")
            success = False
        End Try
        If success = False Then
            success = True
        Else
            Array.Sort(resultsLJump)
            Select Case comp1Score(currEvent)
                Case resultsLJump(0)
                    comp1Points(currEvent) = 1
                Case resultsLJump(1)
                    comp1Points(currEvent) = 2
                Case resultsLJump(2)
                    comp1Points(currEvent) = 3
            End Select
            Select Case comp2Score(currEvent)
                Case resultsLJump(0)
                    comp2Points(currEvent) = 1
                Case resultsLJump(1)
                    comp2Points(currEvent) = 2
                Case resultsLJump(2)
                    comp2Points(currEvent) = 3
            End Select
            Select Case comp3Score(currEvent)
                Case resultsLJump(0)
                    comp3Points(currEvent) = 1
                Case resultsLJump(1)
                    comp3Points(currEvent) = 2
                Case resultsLJump(2)
                    comp3Points(currEvent) = 3
            End Select
            comp1LJumptxt.Enabled = False
            comp2LJumptxt.Enabled = False
            comp3LJumptxt.Enabled = False
            comp1LJumpChk.Enabled = False
            comp2LJumpChk.Enabled = False
            comp3LJumpChk.Enabled = False
        End If
        If ProgressBar1.Value <= 80 Then
            ProgressBar1.Value = ProgressBar1.Value + 20
            overviewTab()
        End If
    End Sub
    Private Sub Comp1LJumpChk_CheckedChanged(sender As Object, e As EventArgs) Handles comp1LJumpChk.CheckedChanged
        If comp1ChkFlag(4) = False Then
            comp1LJumptxt.Enabled = True
            comp1ChkFlag(4) = True
        Else
            comp1LJumptxt.Enabled = False
            comp1ChkFlag(4) = False
        End If
    End Sub
    Private Sub Comp2LJumpChk_CheckedChanged(sender As Object, e As EventArgs) Handles comp2LJumpChk.CheckedChanged
        If comp2ChkFlag(4) = False Then
            comp2LJumptxt.Enabled = True
            comp2ChkFlag(4) = True
        Else
            comp2LJumptxt.Enabled = False
            comp2ChkFlag(4) = False
        End If
    End Sub
    Private Sub Comp3LJumpChk_CheckedChanged(sender As Object, e As EventArgs) Handles comp3LJumpChk.CheckedChanged
        If comp3ChkFlag(4) = False Then
            comp3LJumptxt.Enabled = True
            comp3ChkFlag(4) = True
        Else
            comp3LJumptxt.Enabled = False
            comp3ChkFlag(4) = False
        End If
    End Sub
    Private Sub readFileBtn_Click(sender As Object, e As EventArgs) Handles readFileBtn.Click
        If Not System.IO.File.Exists(strFileName) Then 'if the text file doesnt exist, throw error
            MsgBox("The file does not exist, please create a valid file using the 'Write File' button or move your save-file to your desktop")
        Else 'if the file does exist
            Dim objReader As New System.IO.StreamReader(strFileName) 'read the text file and change all relevant inputs in the program to match the data in the text file.
            junkData = objReader.ReadLine() & vbCrLf 'junk formatting data gets ignored
            compName1txt.Text = objReader.ReadLine()
            compName2txt.Text = objReader.ReadLine()
            compName3txt.Text = objReader.ReadLine()
            eventTabControl.Enabled = True 'enables the tab control, because if you dont, the program can't modify other properties on pages tha tarent currenly being viewed
            junkData = objReader.ReadLine()
            junkData = objReader.ReadLine()
            comp1200mtxt.Text = objReader.ReadLine()
            comp2200mtxt.Text = objReader.ReadLine()
            comp3200mtxt.Text = objReader.ReadLine()
            eventTabControl.SelectTab(1) ' changes the current tab control page because the program cant modify values of screen elements not 'active' in a tab control
            junkData = objReader.ReadLine()
            junkData = objReader.ReadLine()
            comp11500mtxt.Text = objReader.ReadLine()
            comp21500mtxt.Text = objReader.ReadLine()
            comp31500mtxt.Text = objReader.ReadLine()
            eventTabControl.SelectTab(2)
            junkData = objReader.ReadLine()
            junkData = objReader.ReadLine()
            comp1Discustxt.Text = objReader.ReadLine()
            comp2Discustxt.Text = objReader.ReadLine()
            comp3Discustxt.Text = objReader.ReadLine()
            eventTabControl.SelectTab(3)
            junkData = objReader.ReadLine()
            junkData = objReader.ReadLine()
            comp1Javelintxt.Text = objReader.ReadLine()
            comp2Javelintxt.Text = objReader.ReadLine()
            comp3Javelintxt.Text = objReader.ReadLine()
            eventTabControl.SelectTab(4)
            junkData = objReader.ReadLine()
            junkData = objReader.ReadLine()
            comp1LJumptxt.Text = objReader.ReadLine()
            comp2LJumptxt.Text = objReader.ReadLine()
            comp3LJumptxt.Text = objReader.ReadLine()
            objReader.Close() 'stop reading
            confPartBtn.Enabled = True 'enable the confirmation of competitor names button
            eventTabControl.SelectTab(0) 'make the first tab the currently viewed tab
        End If
    End Sub
    Private Sub writeFileBtn_Click(sender As Object, e As EventArgs) Handles writeFileBtn.Click
        If Not System.IO.File.Exists(strFileName) Then 'if the text file doesnt exist:
            System.IO.File.Create(strFileName).Dispose() '... create the text file
        End If
        Dim objWriter As New System.IO.StreamWriter(strFileName) 'startup writing function
        objWriter.WriteLine("Competitors:") 'junk formatting text because user-friendly :)
        objWriter.WriteLine(compName1txt.Text) 'write the inputted score
        objWriter.WriteLine(compName2txt.Text)
        objWriter.WriteLine(compName3txt.Text & vbCrLf) 'skip a line after inputting text for good use of space
        objWriter.WriteLine("200m Results:")
        objWriter.WriteLine(comp1200mtxt.Text)
        objWriter.WriteLine(comp2200mtxt.Text)
        objWriter.WriteLine(comp3200mtxt.Text & vbCrLf)
        objWriter.WriteLine("1500m Results:")
        objWriter.WriteLine(comp11500mtxt.Text)
        objWriter.WriteLine(comp21500mtxt.Text)
        objWriter.WriteLine(comp31500mtxt.Text & vbCrLf)
        objWriter.WriteLine("Discus Results:")
        objWriter.WriteLine(comp1Discustxt.Text)
        objWriter.WriteLine(comp2Discustxt.Text)
        objWriter.WriteLine(comp3Discustxt.Text & vbCrLf)
        objWriter.WriteLine("Javelin Results:")
        objWriter.WriteLine(comp1Javelintxt.Text)
        objWriter.WriteLine(comp2Javelintxt.Text)
        objWriter.WriteLine(comp3Javelintxt.Text & vbCrLf)
        objWriter.WriteLine("Long Jump Results:")
        objWriter.WriteLine(comp1LJumptxt.Text)
        objWriter.WriteLine(comp2LJumptxt.Text)
        objWriter.WriteLine(comp3LJumptxt.Text & vbCrLf)
        objWriter.WriteLine("Overall Placements:")
        objWriter.WriteLine(Comp1PlcLbl.Text)
        objWriter.WriteLine(Comp2PlcLbl.Text)
        objWriter.WriteLine(Comp3PlcLbl.Text)
        objWriter.Close() ' no more writing :|
    End Sub
End Class 'MAX AIREY Y11