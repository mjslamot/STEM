Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs)
        Randomize()

    End Sub


    'SIDE FORM CHANGE START
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles picBio.Click
        TabSideMainD.Hide()
        TabSideBio.Show()
        tmrBiologyMainFont.Enabled = True
        TabsMain.SelectTab(21)

        Call MainLabelReset()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles picPhysics.Click
        TabSideMainD.Hide()
        TabSidePhysics.Show()
        tmrPhysicsMainFont.Enabled = True

        TabsMain.SelectTab(22)

        Call MainLabelReset()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles picMath.Click
        TabSideMainD.Hide()
        TabSideMath.Show()
        tmrMathMainFont.Enabled = True

        TabsMain.SelectTab(23)

        Call MainLabelReset()
    End Sub

    Private Sub picPhysicsInBio_Click(sender As Object, e As EventArgs) Handles picPhysicsInBio.Click
        TabSideBio.Hide()
        TabSidePhysics.Show()
        TabsMain.SelectTab(22)
        tmrPhysicsMainFont.Enabled = True

        Call MainLabelReset()
    End Sub

    Private Sub picMathInBio_Click(sender As Object, e As EventArgs) Handles picMathInBio.Click
        TabSideBio.Hide()
        TabSideMath.Show()
        TabsMain.SelectTab(23)
        tmrMathMainFont.Enabled = True
        Call MainLabelReset()

    End Sub

    Private Sub picPhysicsInMath_Click(sender As Object, e As EventArgs) Handles picPhysicsInMath.Click
        TabSideMath.Hide()
        TabSidePhysics.Show()
        tmrPhysicsMainFont.Enabled = True

        TabsMain.SelectTab(22)
        Call MainLabelReset()
    End Sub

    Private Sub picBioInMath_Click(sender As Object, e As EventArgs) Handles picBioInMath.Click
        TabSideMath.Hide()
        TabSideBio.Show()
        TabsMain.SelectTab(21)
        tmrBiologyMainFont.Enabled = True
        Call MainLabelReset()
    End Sub

    Private Sub picBioInPhysics_Click(sender As Object, e As EventArgs) Handles picBioInPhysics.Click
        TabSidePhysics.Hide()
        TabSideBio.Show()
        TabsMain.SelectTab(21)
        tmrBiologyMainFont.Enabled = True
        Call MainLabelReset()
    End Sub

    Private Sub picMathInPhysics_Click(sender As Object, e As EventArgs) Handles picMathInPhysics.Click
        TabSidePhysics.Hide()
        TabSideMath.Show()
        TabsMain.SelectTab(23)
        tmrMathMainFont.Enabled = True
        Call MainLabelReset()
    End Sub

    Private Sub MainLabelReset()
        inExtra9 = 0
        intExtra10 = 0
        intExtra11 = 0

        bolSize9 = False
        bolSize10 = False
        bolSize11 = False
    End Sub






    'SIDE FORM CHANGE END

    Private Sub lblBeginMain_Click(sender As Object, e As EventArgs) Handles lblBeginMain.Click
        TabMain1.Hide()
        TabMain2.Show()
        picBio.Visible = True
        picMath.Visible = True
        picPhysics.Visible = True
    End Sub
    Private intExtra As Integer
    Private bolSize As Boolean

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If bolSize = False Then
            intExtra = intExtra + 1
            If intExtra > 20 Then
                bolSize = True
            End If
        Else
            intExtra = intExtra - 1
            If intExtra < 0 Then
                bolSize = False
            End If
        End If

        lblBeginMain.Font = New Font("Monotype Corsiva", 72 + intExtra)
    End Sub
    ' Complete

    'MITOSIS!






    ' Mitosis Code beings here

    Private bolInstructionComplete As Boolean
    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        'reset first mitosis
        intCount = 0
        bolSize1 = False
        intExtra1 = 0
        Timer3.Enabled = True
        picCell2.Visible = False
        picCell3.Visible = False
        picCell4.Visible = False
        picCell5.Visible = False
        picCell6.Visible = False
        picCell7.Visible = False
        picCell8.Visible = False
        TabsMain.SelectTab(2)
        Label24.Visible = False
        Label36.Visible = False
        Label37.Visible = False

        'Reset B
        intCount1 = 0
        bolSize2 = False
        intExtra2 = 0

        lblProphase.Visible = True
        lblMeta.Visible = False
        lblAnaphase.Visible = False
        lblTelophase.Visible = False

        picProphase.Visible = True
        picMeta.Visible = False
        picAna.Visible = False
        picTelo.Visible = False

        'Reset C

        bolSize3 = False
        intExtra3 = 0
        bolFlash = False
        Label29.Text = "Here's a small game, to help you test your knowledge on mitosis." & vbCrLf & vbCrLf & "To comeplete this game, you will have 20 seconds to match as many cells as you can to their corresponding mitosis phase."
        lblInstructions.Text = "Instructions"
        intInstructionCount = 0
        TimerMit.Enabled = False

        intPicCount = 0
        intCorrectCount = 0
        intWrong = 0
        intQuestion = 0
        intTime = 0

        bolSize1 = False
        strYesNo1 = ""
        intExtra1 = 0
        lblStartMitosis.Visible = True
        Label34.Text = "Score: 0"
        lblTime.Text = "Time: 20"


    End Sub


    Private intCount As Integer
    Private Sub PictureBox1_Click_1(sender As Object, e As EventArgs) Handles picCell1.Click
        If intCount = 0 Then
            picCell2.Visible = True
            Label24.Text = "Mitosis Count: 1"
            Label24.Visible = True
            intCount = 1

        ElseIf intCount = 1 Then
            picCell3.Visible = True
            picCell4.Visible = True
            Label24.Text = "Mitosis Count: 2"
            intCount = 2

        ElseIf intCount = 2 Then
            picCell5.Visible = True
            picCell6.Visible = True
            picCell7.Visible = True
            picCell8.Visible = True
            Label24.Text = "Mitosis Count: 3"
            Label36.Visible = True
            intCount = 3
        End If
    End Sub

    Private intCountA As Integer


    ' Mitosis Game
    Private btnArray(3) As Button
    Private bolPicArray(8) As Boolean
    Private intPicCount, intCorrectCount, intWrong, intQuestion, intTime As Integer

    ' Mitosis Master Prep
    Private Sub MitosisGame()
        TabsMain.SelectTab(12)
        intTime = 21
        btnArray(0) = btn1
        btnArray(1) = btn2
        btnArray(2) = btn3
        btnArray(3) = btn4
        Timer2.Enabled = True
    End Sub




    ' Reset the Quesiton for Mitosis Master
    Private strYesNo1 As String
    Private Sub Reset()
        Dim intCount, intRandom, intPic As Integer
        Dim bolArray(3) As Boolean

        ' Create a New Question
        intQuestion = intQuestion + 1

        Label34.Text = "Score: " & intCorrectCount
        ' Randomize Buttons
        Do
            Do
                intRandom = 3 * Rnd()
            Loop Until bolArray(intRandom) = False
            If intRandom = 0 Then
                btnArray(intCount).Text = "Prophase"
                btnArray(intCount).Tag = 1
            ElseIf intRandom = 1 Then
                btnArray(intCount).Text = "Metaphase"
                btnArray(intCount).Tag = 2
            ElseIf intRandom = 2 Then
                btnArray(intCount).Text = "Anaphase"
                btnArray(intCount).Tag = 3
            Else
                btnArray(intCount).Text = "Telophase"
                btnArray(intCount).Tag = 4
            End If
            bolArray(intRandom) = True
            intCount = intCount + 1
        Loop Until intCount = 4

        ' Select a Random Picture
        Do
            intPic = 8 * Rnd() + 1
        Loop Until intPic < 9 And intPic >= 1

        If intPic = 1 Then
            PictureBox9.Image = My.Resources.KAV_HD_MF_5MP_Snapshot1_5
            PictureBox9.Tag = 3
        ElseIf intPic = 2 Then
            PictureBox9.Image = My.Resources.KAV_HD_MF_5MP_Snapshot1_8
            PictureBox9.Tag = 3
        ElseIf intPic = 3 Then
            PictureBox9.Image = My.Resources.KAV_HD_MF_5MP_Snapshot2
            PictureBox9.Tag = 2
        ElseIf intPic = 4 Then
            PictureBox9.Image = My.Resources.KAV_HD_MF_5MP_Snapshot2_5
            PictureBox9.Tag = 2
        ElseIf intPic = 5 Then
            PictureBox9.Image = My.Resources.KAV_HD_MF_5MP_Snapshot3
            PictureBox9.Tag = 1
        ElseIf intPic = 6 Then
            PictureBox9.Image = My.Resources.KAV_HD_MF_5MP_Snapshot3_1
            PictureBox9.Tag = 1
        ElseIf intPic = 7 Then
            PictureBox9.Image = My.Resources.KAV_HD_MF_5MP_Snapshot4
            PictureBox9.Tag = 4
        ElseIf intPic = 8 Then
            PictureBox9.Image = My.Resources.KAV_HD_MF_5MP_Snapshot4_1
            PictureBox9.Tag = 4
        End If

        ' 
        btn1.Enabled = True
        btn2.Enabled = True
        btn3.Enabled = True
        btn4.Enabled = True
    End Sub
    ' End of Mitosis Master Reset Question





    ' Code for Buttons (Phases)
    Private Sub btn1_click(sender As Object, e As EventArgs) Handles btn1.Click

        'Check for correct answer
        If PictureBox9.Tag = btn1.Tag Then
            intCorrectCount = intCorrectCount + 1
            Call Reset()

            ' Disabled button, goto Wrong (remove lives)
        Else
            btn1.Enabled = False
            Call Wrong()
        End If
    End Sub
    Private Sub btn2_Click(sender As Object, e As EventArgs) Handles btn2.Click 'Same as btn1
        If PictureBox9.Tag = btn2.Tag Then
            intCorrectCount = intCorrectCount + 1
            Call Reset()
        Else
            btn2.Enabled = False
            Call Wrong()
        End If
    End Sub
    Private Sub btn3_Click(sender As Object, e As EventArgs) Handles btn3.Click ' Same as btn1
        If PictureBox9.Tag = btn3.Tag Then
            intCorrectCount = intCorrectCount + 1
            Call Reset()
        Else
            btn3.Enabled = False
            Call Wrong()
        End If
    End Sub
    Private Sub btn4_Click(sender As Object, e As EventArgs) Handles btn4.Click 'Same as btn1
        If PictureBox9.Tag = btn4.Tag Then
            intCorrectCount = intCorrectCount + 1
            Call Reset()
        Else
            btn4.Enabled = False
            Call Wrong()
        End If
    End Sub
    ' End of Buttons Code

    ' Error/lose of life
    Private Sub Wrong()

        intWrong = intWrong + 1

        ' Take away first two lives
        If intWrong = 1 Then
            PictureBox1.Visible = False
        ElseIf intWrong = 2 Then
            PictureBox4.Visible = False

            ' Take away Final life/Reset
        ElseIf intWrong = 3 Then
            PictureBox5.Visible = False

            Timer1.Enabled = False
            Label1.Text = ""

            strYesNo1 = MsgBox("You ran out of lives :(" & vbCrLf & vbCrLf & "Your Score is: " & intCorrectCount)
            Label29.Text = "Click on 'Begin' to play again or select a new unit"

            Call Replay()
        End If
    End Sub


    ' Countdown
    Private Sub TimerMit_Tick(sender As Object, e As EventArgs) Handles TimerMit.Tick
        intTime = intTime - 1
        lblTime.Text = "Time: " & intTime

        'Times Up!
        If intTime = 0 Then
            TimerMit.Enabled = False

            ' Tell Score
            strYesNo1 = MsgBox("You ran out of time" & vbCrLf & vbCrLf & "Your Score is: " & intCorrectCount)
            Label29.Text = "Click on 'Begin' to play again or select a new unit"
            lblStartMitosis.Visible = True
            Call Replay()

        End If

        ' Temperary Reset/ Not Functioning
        'If strYesNo1 = vbYes Then
        '    Call Reset()
        'ElseIf strYesNo1 = vbNo Then
        'End If

    End Sub

    Private Sub Replay()
        intPicCount = 0
        intCorrectCount = 0
        intWrong = 0
        intQuestion = 0
        intTime = 0
        btn1.Enabled = False
        btn2.Enabled = False
        btn3.Enabled = False
        btn4.Enabled = False

        TimerMit.Enabled = False
        bolSize1 = False
        strYesNo1 = ""
        intExtra1 = 0
        bolFlash = True
        lblStartMitosis.Visible = True
        PictureBox1.Visible = True
        PictureBox4.Visible = True
        PictureBox5.Visible = True

        'Call MitosisGame()
    End Sub
    ' Begins the Game
    Private strSkip As String
    Private Sub lblStart_Click(sender As Object, e As EventArgs) Handles lblStartMitosis.Click
        intTime = 21

        If Not bolInstructionComplete = True Then
            strSkip = MsgBox("Are you sure you want to skip the instructions?", vbYesNo)
            If strSkip = vbYes Then
                bolInstructionComplete = True
                PictureBox9.Visible = True

                btn1.Visible = True
                btn2.Visible = True
                btn3.Visible = True
                btn4.Visible = True
                btn1.Enabled = True
                btn2.Enabled = True
                btn3.Enabled = True
                btn4.Enabled = True

                PictureBox1.Visible = True
                PictureBox4.Visible = True
                PictureBox5.Visible = True
                Label34.Visible = True
                lblTime.Visible = True

                bolFlash = True
                lblStartMitosis.Enabled = True


                GoTo Begin
            Else


            End If
        Else

Begin:
            Label29.Text = "Have Fun!"
            lblInstructions.Visible = False
            lblStartMitosis.Visible = False
            btn1.Enabled = True
            btn2.Enabled = True
            btn3.Enabled = True
            btn4.Enabled = True

            ' Countdown Timer Enabling
            TimerMit.Enabled = True
            Call Reset()
        End If
    End Sub




    Private bolSize3 As Boolean
    Private intExtra3 As Integer

    ' Word Flash Timer
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        If bolSize3 = False Then
            intExtra3 = intExtra3 + 1

            ' Goes up 20, then start going back down
            If intExtra3 > 20 Then
                bolSize3 = True
            End If
        Else

            ' Goes down to 0 and then goes back up
            intExtra3 = intExtra3 - 1
            If intExtra3 < 0 Then
                bolSize3 = False
            End If
        End If

        ' Instructions Flash
        If bolFlash = False Then
            lblInstructions.Font = New Font("Monotype Corsiva", 36 + intExtra3)
        Else
            ' Begin Button Flash
            lblStartMitosis.Font = New Font("Monotype Corsiva", 36 + intExtra3)
        End If
    End Sub

    ' Instructions
    Dim bolFlash As Boolean
    Dim intInstructionCount As Integer
    Private Sub Label30_Click(sender As Object, e As EventArgs) Handles lblInstructions.Click

        ' Step 1: Pictures
        If intInstructionCount = 0 Then
            Label29.Text = "Cells will appear in the center of the screen as so. However, they will be pictures of real cells as seen under high-power microscopes"
            lblInstructions.Text = "Continue"
            PictureBox9.Visible = True

            ' Step 2: Buttons
        ElseIf intInstructionCount = 1 Then
            Label29.Text = "Your objective is to select the button that corresponds to the image's phase"
            btn1.Visible = True
            btn2.Visible = True
            btn3.Visible = True
            btn4.Visible = True

            ' Step 3: Lives, Score, Time
        ElseIf intInstructionCount = 2 Then
            Label29.Text = "Your score, your lives, and your time remaining will be kept to the right"
            PictureBox1.Visible = True
            PictureBox4.Visible = True
            PictureBox5.Visible = True
            Label34.Visible = True
            lblTime.Visible = True

            ' Begin
        ElseIf intInstructionCount = 3 Then
            Label29.Text = "Click on 'Begin' when you are ready"
            lblInstructions.Visible = False
            bolInstructionComplete = True
            bolFlash = True
        End If

        ' Add 1 Step
        intInstructionCount = intInstructionCount + 1
    End Sub


    ' Reset for Mitosis slide 1
    Private Sub Label36_Click(sender As Object, e As EventArgs) Handles Label36.Click
        Label24.Visible = False
        Label36.Visible = False
        picCell3.Visible = False
        picCell2.Visible = False
        picCell4.Visible = False
        picCell5.Visible = False
        picCell6.Visible = False
        picCell7.Visible = False
        picCell8.Visible = False
        intCount = 0
        bolSize1 = False
        intExtra1 = 0
        Timer3.Enabled = False
        Timer4.Enabled = True
        TabsMain.SelectTab(3)
    End Sub

    Private bolSize1 As Boolean
    Private intExtra1 As Integer
    ' Mitosis slide 1 flashing
    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick

        If bolSize1 = False Then
            intExtra1 = intExtra1 + 1

            ' Goes up 20, then start going back down
            If intExtra1 > 10 Then
                bolSize1 = True
            End If
        Else

            ' Goes down to 0 and then goes back up
            intExtra1 = intExtra1 - 1
            If intExtra1 < 0 Then
                bolSize1 = False
            End If
        End If

        ' Instructions Flash
        If Not intCount = 3 Then
            picCell1.Width = 223 + intExtra1
            picCell1.Height = 152 + intExtra
        Else
            ' Begin Button Flash
            Label36.Font = New Font("Kristen ITC", 36 + intExtra1)
        End If

    End Sub


    Private Sub Label37_Click(sender As Object, e As EventArgs) Handles Label37.Click


        intCountA = 0
        lblTelophase.Visible = False
        picTelo.Visible = False
        lblProphase.Visible = True
        picProphase.Visible = True
        Call MitosisGame()
        ' Initiates the Mitosis Game

    End Sub

    Private bolSize2 As Boolean
    Private intExtra2 As Integer
    Private intCount1 As Integer
    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        If bolSize2 = False Then
            intExtra2 = intExtra2 + 1

            If intExtra2 > 10 Then
                bolSize2 = True
            End If
        Else
            intExtra2 = intExtra2 - 1
            If intExtra2 < 0 Then
                bolSize2 = False
            End If
        End If


        Select Case intCount1
            Case 0
                PictureBox2.Width = 208 + intExtra2
                PictureBox2.Height = 141 + intExtra2
            Case 1
                PictureBox3.Width = 208 + intExtra2
                PictureBox3.Height = 141 + intExtra2
            Case 2
                PictureBox6.Width = 208 + intExtra2
                PictureBox6.Height = 141 + intExtra2
            Case 3
                PictureBox7.Width = 208 + intExtra2
                PictureBox7.Height = 141 + intExtra2
            Case 4
                Label37.Font = New Font("Kristen ITC", 36 + intExtra2)
        End Select
    End Sub

    Private Sub PictureBox2_Click_1(sender As Object, e As EventArgs) Handles PictureBox2.Click
        intCount1 = 0
        lblProphase.Visible = True
        picProphase.Visible = True
        lblMeta.Visible = False
        picMeta.Visible = False
        lblAnaphase.Visible = False
        picAna.Visible = False
        lblTelophase.Visible = False
        picTelo.Visible = False
    End Sub

    Private Sub PictureBox3_Click_1(sender As Object, e As EventArgs) Handles PictureBox3.Click
        intCount1 = 1
        lblProphase.Visible = False
        picProphase.Visible = False
        lblAnaphase.Visible = False
        picAna.Visible = False
        lblTelophase.Visible = False
        picTelo.Visible = False
        lblMeta.Visible = True
        picMeta.Visible = True
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        intCount1 = 2
        lblProphase.Visible = False
        picProphase.Visible = False
        lblMeta.Visible = False
        picMeta.Visible = False
        lblTelophase.Visible = False
        picTelo.Visible = False
        lblAnaphase.Visible = True
        picAna.Visible = True
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        intCount1 = 3
        lblProphase.Visible = False
        picProphase.Visible = False
        lblMeta.Visible = False
        picMeta.Visible = False
        lblAnaphase.Visible = False
        picAna.Visible = False
        lblTelophase.Visible = True
        picTelo.Visible = True
        Label37.Visible = True
    End Sub




    ' Kinematics
    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        TabsMain.SelectTab(14)
        Timer7.Enabled = True
        txtAnswer1.Text = ""
        Randomize()

        'Reset Code
        Label49.Visible = True
        Label51.Visible = False
        txtAnswer1.Visible = False
        btnAns.Visible = False
        Label50.Visible = False
        Label48.Visible = False
        Label47.Visible = False
        Label46.Text = "Solve the following questions using your kinematic formulas. When you feel ready, hide the formulas."
        bolSize5 = False
        intExtra5 = 0
        strName1 = ""
        strAction1 = ""
        intDistance = 0
        intTime1 = 0
        intVelocity2 = 0
        intName = 0
        intAction = 0
        intQtype = 0
        intAcc = 0
        intAnswer2 = 0
        intVelocity = 0
        intAnswer1 = 0
    End Sub



    Private bolSize5 As Boolean
    Private intExtra5 As Integer

    Private Sub Timer6_Tick(sender As Object, e As EventArgs) Handles Timer6.Tick
        If bolSize5 = False Then
            intExtra5 = intExtra5 + 1

            If intExtra5 > 10 Then
                bolSize5 = True
            End If
        Else
            intExtra5 = intExtra5 - 1
            If intExtra5 < 0 Then
                bolSize5 = False
            End If
        End If

        Label41.Font = New Font("Kristen ITC", 36 + intExtra5)
    End Sub

    Private Sub Label41_Click(sender As Object, e As EventArgs) Handles Label41.Click
        TabsMain.SelectTab(14)
        Timer7.Enabled = True
    End Sub

    Private Sub Label48_Click(sender As Object, e As EventArgs) Handles Label48.Click
        If Label47.Visible = True Then
            Label47.Visible = False
            Label48.Text = "Show Formulas"
        Else
            Label48.Text = "Hide Formulas"
            Label47.Visible = True
        End If
    End Sub



    Private strName1, strAction1 As String
    Private intDistance, intTime1, intVelocity2, intName, intAction, intQtype, intAcc, intAnswer2 As Integer

    Private Sub Label49_Click(sender As Object, e As EventArgs) Handles Label49.Click
        Label51.Visible = True
        txtAnswer1.Visible = True
        btnAns.Visible = True
        Label47.Visible = True
        Label48.Visible = True
        Label49.Visible = False
        Call ResetKine()

    End Sub

    Private intAnswer1 As Integer
    Private intVelocity As Integer



    Private Sub ResetKine()
        Randomize()

        Label50.Visible = True
        intName = 3 * Rnd() + 1
        intAction = 3 * Rnd() + 1
        intQtype = 3 * Rnd() + 1
        Select Case intName
            Case 1
                strName1 = "cat"
            Case 2
                strName1 = "dog"
            Case Else
                strName1 = "frog"
        End Select
        Select Case intAction
            Case 1
                strAction1 = "travels"
            Case 2
                strAction1 = "runs"
            Case Else
                strAction1 = "walks"
        End Select

        Select Case intQtype
            Case 1
                Do
                    intDistance = 1000 * Rnd() + 1
                    intTime1 = 50 * Rnd() + 1
                Loop Until intDistance > intTime1 And intDistance Mod intTime1 = 0
                intAnswer1 = intDistance / intTime1
                Label50.Text = "Meters per Hour"

            Case 2
                Do
                    intVelocity = 1000 * Rnd() + 1
                    intTime1 = 50 * Rnd() + 1
                Loop Until intVelocity > 1
                intAnswer1 = intVelocity * intTime1
                Label50.Text = "Meters"

            Case 3
                Do
                    intVelocity = 100 * Rnd() + 1
                    intDistance = 1000 * Rnd() + 1
                Loop Until intDistance > intVelocity And intDistance Mod intVelocity = 0 And intVelocity > 1
                intAnswer1 = intDistance / intVelocity
                Label50.Text = "Hours"

        End Select

        If intQtype = 1 Then
            Label46.Text = "A" & Space(1) & strName1 & Space(1) & strAction1 & Space(1) & intDistance & Space(1) & "metres in" & Space(1) & intTime1 & Space(1) & "hours. Find the speed at which the" & Space(1) & strName1 & Space(1) & "is moving."
        ElseIf intQtype = 2 Then
            Label46.Text = "A" & Space(1) & strName1 & Space(1) & strAction1 & Space(1) & "at a speed of" & Space(1) & intVelocity & Space(1) & "metres per hour. How far does it travel in" & Space(1) & intTime1 & Space(1) & "hours?"
        ElseIf intQtype = 3 Then
            Label46.Text = "A" & Space(1) & strName1 & Space(1) & strAction1 & Space(1) & "at a speed of" & Space(1) & intVelocity & Space(1) & "metres per hour. How long does it take for it to travel" & Space(1) & intDistance & Space(1) & "metres?"
        End If
    End Sub

    Private intErrorCountPhysics As Integer
    Private Sub btnAns_Click(sender As Object, e As EventArgs) Handles btnAns.Click
        Dim strTest As String
        strTest = txtAnswer1.Text

        If IsNumeric(strTest) = True Then
            intAnswer2 = txtAnswer1.Text

            If intAnswer1 = intAnswer2 Then
                MsgBox("You got it right!")
                txtAnswer1.Text = ""
                intErrorCountPhysics = 0
                Call ResetKine()
            Else

                intErrorCountPhysics = intErrorCountPhysics + 1
                If intErrorCountPhysics < 3 Then
                    MsgBox("Wrong! Try Again")
                Else
                    MsgBox("Wrong! The correct answer was: " & intAnswer1)
                    intErrorCountPhysics = 0
                    txtAnswer1.Text = ""
                    Call ResetKine()
                End If
            End If
        Else
            If txtAnswer1.Text = "" Then
                MsgBox("Please enter your answer")
            Else
                MsgBox("Please only enter digits")
                txtAnswer1.Text = ""
                Call ResetKine()
            End If
        End If
    End Sub



















    ' Bio DNA MATCHUP

    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click
        TabsMain.SelectTab(15)
        Timer8.Enabled = True
        Call DNAReset()
    End Sub


    Private bolSize6 As Boolean
    Private intExtra6 As Integer
    Private Sub Timer8_Tick(sender As Object, e As EventArgs) Handles Timer8.Tick
        If bolSize6 = False Then
            intExtra6 = intExtra6 + 1

            If intExtra6 > 10 Then
                bolSize6 = True
            End If
        Else
            intExtra6 = intExtra6 - 1
            If intExtra6 < 0 Then
                bolSize6 = False
            End If
        End If

        Label55.Font = New Font("Kristen ITC", 36 + intExtra6)
    End Sub



    Private Sub Label55_Click(sender As Object, e As EventArgs) Handles Label55.Click
        TabsMain.SelectTab(16)
        Timer8.Enabled = False
        Timer9.Enabled = True
    End Sub

    'Game begisn here
    Private intExtra7 As Integer
    Private bolSize7 As Boolean
    Private bolControlsDone As Boolean

    Private Sub Timer9_Tick(sender As Object, e As EventArgs) Handles Timer9.Tick
        If bolSize7 = False Then
            intExtra7 = intExtra7 + 1

            If intExtra7 > 10 Then
                bolSize7 = True
            End If
        Else
            intExtra7 = intExtra7 - 1
            If intExtra7 < 0 Then
                bolSize7 = False
            End If
        End If

        lblGuideForDNAGuru.Font = New Font("Kristen ITC", 36 + intExtra7)
    End Sub

    Private Sub lblGuideForDNAGuru_Click(sender As Object, e As EventArgs) Handles lblGuideForDNAGuru.Click
        If bolControlsDone = False Then
            Label59.Visible = True
            lblGuideForDNAGuru.Text = "Begin"
            bolControlsDone = True
        Else
            lblGuideForDNAGuru.Visible = False
            Call DNALoad()
            tmrDNAGame.Enabled = True
            Label58.Visible = True
            Label60.Visible = True
            Label58.Text = "Score: 0"
            Label60.Text = "Level: 1"
        End If
    End Sub





    Private lblArray(6) As Label



    Private intSpeed, intDNAScore, intDNATickCount, intDNARandom1 As Integer
    Private bolOn As Boolean

    Private Sub DNALoad()
        Dim intDNACountB As Integer

        Randomize()

        intSpeed = 4

        lblArray(0) = lblDNA1
        lblArray(1) = lblDNA2
        lblArray(2) = lblDNA3
        lblArray(3) = lblDNA4
        lblArray(4) = lblDNA5
        lblArray(5) = lblDNA6
        lblArray(6) = lblDNA7

        Do
            Do
                intDNARandom1 = 4 * Rnd() + 1
            Loop Until intDNARandom1 < 5 And intDNARandom1 > 0

            If intDNARandom1 = 1 Then
                lblArray(intDNACountB).Text = "A"
            ElseIf intDNARandom1 = 2 Then
                lblArray(intDNACountB).Text = "T"
            ElseIf intDNARandom1 = 3 Then
                lblArray(intDNACountB).Text = "C"
            ElseIf intDNARandom1 = 4 Then
                lblArray(intDNACountB).Text = "G"
            End If
            intDNACountB = intDNACountB + 1
        Loop Until intDNACountB = 5

        tmrDNAGame.Enabled = True

        bolOn = True
    End Sub




    Private Sub tmrDNAGame_Tick(sender As Object, e As EventArgs) Handles tmrDNAGame.Tick
        Dim intCountDNA1 As Integer

        Do
            lblArray(intCountDNA1).Left -= intSpeed
            intCountDNA1 = intCountDNA1 + 1
        Loop Until intCountDNA1 = 7

        intCountDNA1 = 0
        '
        Do
            If lblArray(intCountDNA1).Location.X <= 0 Then
                tmrDNAGame.Enabled = False
                MsgBox("Good Job! Your score is:" & Space(2) & intDNAScore)

                Call DNAReset

                ' Reset Goes Here
            End If
            intCountDNA1 = intCountDNA1 + 1
        Loop Until intCountDNA1 = 7
        intCountDNA1 = 0

        intDNATickCount = intDNATickCount + 1


    End Sub



    ' Fix
    Private intDNACountC As Integer
    Private Sub button1_Keyodown(Sender As Object, e As KeyEventArgs) Handles Button1.KeyDown
        Dim intDNARandomNumber2 As Integer
        Dim bolLock As Boolean

        If bolOn = True Then


            If e.KeyCode = Keys.A Then
                If lblArray(intDNACountC).Text = "T" Then
                    GoTo DNACorrect
                Else
                    GoTo DNAWrong
                End If
            End If

            If e.KeyCode = Keys.T Then
                If lblArray(intDNACountC).Text = "A" Then
                    GoTo DNACorrect
                Else
                    GoTo DNAWrong
                End If
            End If

            If e.KeyCode = Keys.G Then
                If lblArray(intDNACountC).Text = "C" Then
                    GoTo DNACorrect
                Else
                    GoTo DNAWrong
                End If
            End If

            If e.KeyCode = Keys.C Then
                If lblArray(intDNACountC).Text = "G" Then
                    GoTo DNACorrect
                Else
                    GoTo DNAWrong
                End If
            End If

            If bolLock = True Then
DNACorrect:
                intDNAScore = intDNAScore + 1

                If intDNACountC = 0 Then
                    lblArray(0).Left = lblArray(6).Location.X + 200
                ElseIf intDNACountC = 1 Then
                    lblArray(1).Left = lblArray(0).Location.X + 200
                ElseIf intDNACountC = 2 Then
                    lblArray(2).Left = lblArray(1).Location.X + 200
                ElseIf intDNACountC = 3 Then
                    lblArray(3).Left = lblArray(2).Location.X + 200
                ElseIf intDNACountC = 4 Then
                    lblArray(4).Left = lblArray(3).Location.X + 200
                ElseIf intDNACountC = 5 Then
                    lblArray(5).Left = lblArray(4).Location.X + 200
                ElseIf intDNACountC = 6 Then
                    lblArray(6).Left = lblArray(5).Location.X + 200
                End If

                Do
                    intDNARandomNumber2 = 4 * Rnd() + 1
                Loop Until intDNARandomNumber2 <> 0
                If intDNARandomNumber2 = 1 Then
                    lblArray(intDNACountC).Text = "A"
                ElseIf intDNARandomNumber2 = 2 Then
                    lblArray(intDNACountC).Text = "T"
                ElseIf intDNARandomNumber2 = 3 Then
                    lblArray(intDNACountC).Text = "C"
                ElseIf intDNARandomNumber2 = 4 Then
                    lblArray(intDNACountC).Text = "G"
                End If

                If intDNACountC = 6 Then
                    intDNACountC = 0
                Else
                    intDNACountC = intDNACountC + 1
                End If
            End If

            If bolLock = True Then
DNAWrong:
                tmrDNAGame.Enabled = False
                MsgBox("Good Job! Your score is:" & Space(2) & intDNAScore)

                Call DNAReset
                ' Reset
            End If
        End If

        Label58.Text = "Score: " & intDNAScore
        If intDNAScore Mod 10 = 0 And intDNAScore <> 0 Then
            intSpeed = intSpeed + 1
            Label60.Text = "Level: " & intSpeed - 3
        End If




    End Sub


    Private Sub DNAReset()
        If bolControlsDone = True Then
            lblGuideForDNAGuru.Text = "Begin"
        Else
            lblGuideForDNAGuru.Text = "Controls"
        End If

        lblGuideForDNAGuru.Visible = True
        intDNACountC = 0
        intSpeed = 4
        intDNAScore = 0
        intDNATickCount = 0
        intDNARandom1 = 0
        bolOn = False
        Label60.Text = "Level: 1"
        tmrDNAGame.Enabled = False

        lblDNA1.Location = New Point(1300, 475)
        lblDNA2.Location = New Point(1500, 475)
        lblDNA3.Location = New Point(1700, 475)
        lblDNA4.Location = New Point(1900, 475)
        lblDNA5.Location = New Point(2100, 475)
        lblDNA6.Location = New Point(2300, 475)
        lblDNA7.Location = New Point(2500, 475)
    End Sub













    ' Cell Theory Begins here
    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        Timer5.Enabled = True
        TabsMain.SelectTab(8)
        Call OrganelleReset()

    End Sub

    Private bolsize4 As Boolean
    Private intExtra4 As Integer
    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick
        If bolsize4 = False Then
            intExtra4 = intExtra4 + 1

            If intExtra4 > 10 Then
                bolsize4 = True
            End If
        Else
            intExtra4 = intExtra4 - 1
            If intExtra4 < 0 Then
                bolsize4 = False
            End If
        End If

        Label38.Font = New Font("Kristen ITC", 36 + intExtra4)
    End Sub

    Private Sub Label38_Click(sender As Object, e As EventArgs) Handles Label38.Click
        TabsMain.SelectTab(17)
        Timer5.Enabled = False
        Timer10.Enabled = True
    End Sub

    Private bolsize8 As Boolean
    Private intExtra8 As Boolean
    Private Sub Timer10_Tick(sender As Object, e As EventArgs) Handles Timer10.Tick
        If bolsize8 = False Then
            intExtra8 = intExtra8 + 1

            If intExtra8 > 10 Then
                bolsize8 = True
            End If
        Else
            intExtra8 = intExtra8 - 1
            If intExtra8 < 0 Then
                bolsize8 = False
            End If
        End If

        Label63.Font = New Font("Kristen ITC", 36 + intExtra8)
    End Sub

    Private Sub Label63_Click(sender As Object, e As EventArgs) Handles Label63.Click
        Dim intCountP As Integer

        Label63.Visible = False
        tmrOrganelle.Enabled = True
        GroupBox1.Visible = True

        lblArrayOrganelle(0) = lblNucleolus
        lblArrayOrganelle(2) = lblChloroplast
        lblArrayOrganelle(1) = lblChromosomes
        lblArrayOrganelle(3) = lblCentriole
        lblArrayOrganelle(4) = lblCellWall
        lblArrayOrganelle(5) = lblEndoplasm
        lblArrayOrganelle(6) = lblLysosome
        lblArrayOrganelle(7) = lblCellMembrane
        lblArrayOrganelle(8) = lblRibosome
        lblArrayOrganelle(9) = lblNucleus
        lblArrayOrganelle(10) = lblMitochrondrion
        lblArrayOrganelle(11) = lblCytoplasm
        lblArrayOrganelle(12) = lblVacuole
        lblArrayOrganelle(13) = lblGolgiApparatus

        Do
            lblArrayOrganelle(intCountP).Visible = True
            intCountP = intCountP + 1
        Loop Until intCountP = 14



        Call PreSetUp
    End Sub


    ' Game Begins, Choses button directions
    Private lblArrayOrganelle(13) As Label
    Private intInitialDirectionArray(13, 1) As Integer
    Private Sub PreSetUp()
        Dim intOrganelleCount As Integer
        Randomize()

        Do
            Do
                intInitialDirectionArray(intOrganelleCount, 0) = 1 * Rnd() + 1
                intInitialDirectionArray(intOrganelleCount, 1) = 1 * Rnd() + 1
            Loop Until intInitialDirectionArray(intOrganelleCount, 0) = 1 Or intInitialDirectionArray(intOrganelleCount, 0) = 2 And intInitialDirectionArray(intOrganelleCount, 1) = 1 Or intInitialDirectionArray(intOrganelleCount, 0) = 2
            If intInitialDirectionArray(intOrganelleCount, 0) = 2 Then
                intInitialDirectionArray(intOrganelleCount, 0) = -1
            End If
            If intInitialDirectionArray(intOrganelleCount, 1) = 2 Then
                intInitialDirectionArray(intOrganelleCount, 1) = -1
            End If
            intOrganelleCount = intOrganelleCount + 1
        Loop Until intOrganelleCount = 14
        intOrganelleCount = 0

        Do
            lblArrayOrganelle(intOrganelleCount).Tag = intOrganelleCount
            intOrganelleCount = intOrganelleCount + 1
        Loop Until intOrganelleCount = 14

        Call QSetUp()
    End Sub


    ' Choses Question
    Private intOrganelleCorrectCounter As Integer
    Private bolOrganelleArray(13) As Boolean
    Private Sub QSetUp()
        Dim intOrganelleRandom As Integer

        If intOrganelleCorrectCounter = 14 Then
            Call OrganelleReset()

            ' Reset Code
        Else
            Do
                Do
                    intOrganelleRandom = 14 * Rnd()
                Loop Until intOrganelleRandom < 14 And intOrganelleRandom >= 0
            Loop Until bolOrganelleArray(intOrganelleRandom) = False
            If intOrganelleRandom = 0 Then
                Label62.Text = "Spherical structure within the nucleus of cells. It is involved in the making of protein"
                Label62.Tag = 0
            ElseIf intOrganelleRandom = 1 Then
                Label62.Text = "Where genetic information is organised into thread-like structures"
                Label62.Tag = 1
            ElseIf intOrganelleRandom = 2 Then
                Label62.Text = "The site of photosynthesis in plants."
                Label62.Tag = 2
            ElseIf intOrganelleRandom = 3 Then
                Label62.Text = "Small protein structure critical to cell division found only in animal cells"
                Label62.Tag = 3
            ElseIf intOrganelleRandom = 4 Then
                Label62.Text = "Structure made of cellulose that protects and supports plant cells"
                Label62.Tag = 4
            ElseIf intOrganelleRandom = 5 Then
                Label62.Text = "Series of canal that carry materials throughout the cell"
                Label62.Tag = 5
            ElseIf intOrganelleRandom = 6 Then
                Label62.Text = "Sac-like structure that contain proteins which break down large molecules"
                Label62.Tag = 6
            ElseIf intOrganelleRandom = 7 Then
                Label62.Text = "Acts like a gate-keeper. Controls movement of materials into and out of the cell"
                Label62.Tag = 7
            ElseIf intOrganelleRandom = 8 Then
                Label62.Text = "Builds proteins essential for cell growth and reproduction"
                Label62.Tag = 8
            ElseIf intOrganelleRandom = 9 Then
                Label62.Text = "The control centre of the cell, directs all of the cell's activities."
                Label62.Tag = 9
            ElseIf intOrganelleRandom = 10 Then
                Label62.Text = "Tiny oval-shaped organelle that provides cells with energy."
                Label62.Tag = 10
            ElseIf intOrganelleRandom = 11 Then
                Label62.Text = "Jelly-like fluid in the cell where nutrients are absorbed, transported and processed."
                Label62.Tag = 11
            ElseIf intOrganelleRandom = 12 Then
                Label62.Text = "Fluid filled space containing water, sugar, minerals and proteins."
                Label62.Tag = 12
            ElseIf intOrganelleRandom = 13 Then
                Label62.Text = "A structure that stores proteins until needed for use inside or outside the cell."
                Label62.Tag = 13
            End If
            bolOrganelleArray(intOrganelleRandom) = True
        End If
    End Sub



    Private Sub tmrOrganelle_Tick(sender As Object, e As EventArgs) Handles tmrOrganelle.Tick
        Dim intCountOrganelle As Integer
        Dim intTop, intRight As Integer


        intTop = GroupBox1.Height - lblNucleolus.Height
        intRight = GroupBox1.Width - lblNucleolus.Width

        Do
            lblArrayOrganelle(intCountOrganelle).Left -= intInitialDirectionArray(intCountOrganelle, 0)
            lblArrayOrganelle(intCountOrganelle).Top -= intInitialDirectionArray(intCountOrganelle, 1)
            intCountOrganelle = intCountOrganelle + 1
        Loop Until intCountOrganelle = 14
        intCountOrganelle = 0


        ' Where the magic happens WIP
        Dim intMagicCount As Integer
        Dim intWidthArray(13) As Integer

        intWidthArray(0) = 153
        intWidthArray(1) = 250
        intWidthArray(2) = 194
        intWidthArray(3) = 153
        intWidthArray(4) = 134
        intWidthArray(5) = 325
        intWidthArray(6) = 144
        intWidthArray(7) = 215
        intWidthArray(8) = 147
        intWidthArray(9) = 125
        intWidthArray(10) = 220
        intWidthArray(11) = 154
        intWidthArray(12) = 124
        intWidthArray(13) = 236


        Do
            If lblArrayOrganelle(intCountOrganelle).Top > intTop Then
                intInitialDirectionArray(intCountOrganelle, 1) = 1
            End If
            If lblArrayOrganelle(intCountOrganelle).Top < 5 Then
                intInitialDirectionArray(intCountOrganelle, 1) = -1
            End If
            If lblArrayOrganelle(intCountOrganelle).Left < 0 Then
                intInitialDirectionArray(intCountOrganelle, 0) = -1
            End If
            If lblArrayOrganelle(intCountOrganelle).Left + intWidthArray(intCountOrganelle) > intRight + 160 Then
                intInitialDirectionArray(intCountOrganelle, 0) = 1
            End If
            intCountOrganelle = intCountOrganelle + 1
        Loop Until intCountOrganelle = 14





       

        ' almost works, try to use > and < ranges to make it work perfectly
        Do
            If lblCellWall.Location.X + 134 = lblArrayOrganelle(intMagicCount).Location.X Or lblCellWall.Location.X + 133 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblCellWall.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblCellWall.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblCellWall.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If
            If lblCellWall.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblCellWall.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblCellWall.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblCellWall.Location.X + 134 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblCellWall.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0



        Do
            If lblGolgiApparatus.Location.X + 236 = lblArrayOrganelle(intMagicCount).Location.X Or lblGolgiApparatus.Location.X + 237 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblGolgiApparatus.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblGolgiApparatus.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblGolgiApparatus.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If


            If lblGolgiApparatus.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblGolgiApparatus.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblGolgiApparatus.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblGolgiApparatus.Location.X + 236 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblGolgiApparatus.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0


        Do
            If lblMitochrondrion.Location.X + 220 = lblArrayOrganelle(intMagicCount).Location.X Or lblMitochrondrion.Location.X + 220 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblMitochrondrion.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblMitochrondrion.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblMitochrondrion.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If
            If lblMitochrondrion.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblMitochrondrion.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblMitochrondrion.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblMitochrondrion.Location.X + 220 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblMitochrondrion.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0


        Do
            If lblEndoplasm.Location.X + 325 = lblArrayOrganelle(intMagicCount).Location.X Or lblEndoplasm.Location.X + 325 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblEndoplasm.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblEndoplasm.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblEndoplasm.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If
            If lblEndoplasm.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblEndoplasm.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblEndoplasm.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblEndoplasm.Location.X + 325 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblEndoplasm.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0


        Do
            If lblCellMembrane.Location.X + 215 = lblArrayOrganelle(intMagicCount).Location.X Or lblCellMembrane.Location.X + 215 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblCellMembrane.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblCellMembrane.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblCellMembrane.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If
            If lblCellMembrane.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblCellMembrane.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblCellMembrane.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblCellMembrane.Location.X + 215 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblCellMembrane.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0




        Do
            If lblVacuole.Location.X + 124 = lblArrayOrganelle(intMagicCount).Location.X Or lblVacuole.Location.X + 124 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblVacuole.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblVacuole.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblVacuole.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If
            If lblVacuole.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblVacuole.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblVacuole.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblVacuole.Location.X + 124 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblVacuole.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0





        Do
            If lblLysosome.Location.X + 144 = lblArrayOrganelle(intMagicCount).Location.X Or lblLysosome.Location.X + 144 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblLysosome.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblLysosome.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblLysosome.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If
            If lblLysosome.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblLysosome.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblLysosome.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblLysosome.Location.X + 144 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblLysosome.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0




        Do
            If lblCytoplasm.Location.X + 154 = lblArrayOrganelle(intMagicCount).Location.X Or lblCytoplasm.Location.X + 154 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblCytoplasm.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblCytoplasm.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblCytoplasm.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If
            If lblCytoplasm.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblCytoplasm.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblCytoplasm.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblCytoplasm.Location.X + 154 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblCytoplasm.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0




        Do
            If lblNucleolus.Location.X + 153 = lblArrayOrganelle(intMagicCount).Location.X Or lblNucleolus.Location.X + 153 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblNucleolus.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblNucleolus.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblNucleolus.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If
            If lblNucleolus.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblNucleolus.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblNucleolus.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblNucleolus.Location.X + 153 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblNucleolus.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0



        Do
            If lblChromosomes.Location.X + 250 = lblArrayOrganelle(intMagicCount).Location.X Or lblChromosomes.Location.X + 250 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblChromosomes.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblChromosomes.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblChromosomes.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If
            If lblChromosomes.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblChromosomes.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblChromosomes.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblChromosomes.Location.X + 250 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblChromosomes.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0






        Do
            If lblCentriole.Location.X + 153 = lblArrayOrganelle(intMagicCount).Location.X Or lblCentriole.Location.X + 153 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblCentriole.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblCentriole.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblCentriole.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If
            If lblCentriole.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblCentriole.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblCentriole.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblCentriole.Location.X + 153 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblCentriole.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0




        Do
            If lblRibosome.Location.X + 147 = lblArrayOrganelle(intMagicCount).Location.X Or lblRibosome.Location.X + 147 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblRibosome.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblRibosome.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblRibosome.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If
            If lblRibosome.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblRibosome.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblRibosome.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblRibosome.Location.X + 147 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblRibosome.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0




        Do
            If lblChloroplast.Location.X + 194 = lblArrayOrganelle(intMagicCount).Location.X Or lblChloroplast.Location.X + 194 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblChloroplast.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblChloroplast.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblChloroplast.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If
            If lblChloroplast.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblChloroplast.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblChloroplast.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblChloroplast.Location.X + 194 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblChloroplast.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0




        Do
            If lblNucleus.Location.X + 125 = lblArrayOrganelle(intMagicCount).Location.X Or lblNucleus.Location.X + 125 = lblArrayOrganelle(intMagicCount).Location.X Then
                If lblNucleus.Location.Y + 36 >= lblArrayOrganelle(intMagicCount).Location.Y And lblNucleus.Location.Y <= lblArrayOrganelle(intMagicCount).Location.X + 36 Then
                    intInitialDirectionArray(lblNucleus.Tag, 0) = 1
                    intInitialDirectionArray(intMagicCount, 0) = -1
                End If
            End If
            If lblNucleus.Location.Y + 36 = lblArrayOrganelle(intMagicCount).Location.Y Or lblNucleus.Location.Y + 35 = lblArrayOrganelle(intMagicCount).Location.Y Then
                tmrOrganelle.Enabled = True
                If lblNucleus.Location.X <= lblArrayOrganelle(intMagicCount).Location.X + intWidthArray(intMagicCount) And lblNucleus.Location.X + 125 >= lblArrayOrganelle(intMagicCount).Location.X Then
                    intInitialDirectionArray(lblNucleus.Tag, 1) = 1
                    intInitialDirectionArray(intMagicCount, 1) = -1
                End If
            End If
            intMagicCount = intMagicCount + 1
        Loop Until intMagicCount = 14
        intMagicCount = 0

    End Sub











    Private Sub lblNucleus_CLick(sender As Object, e As EventArgs) Handles lblNucleus.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblNucleus.Tag Then
            lblNucleus.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()

        Else
            lblNucleus.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub




    Private Sub lblVacuole_click(sender As Object, e As EventArgs) Handles lblVacuole.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblVacuole.Tag Then
            lblVacuole.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()

        Else
            lblVacuole.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click
        TabsMain.SelectTab(18)
    End Sub

    Private Sub Label30_Click_1(sender As Object, e As EventArgs) Handles Label30.Click
        TabsMain.SelectTab(4)
    End Sub

    Private Sub lblNucleolus_click(sender As Object, e As EventArgs) Handles lblNucleolus.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblNucleolus.Tag Then
            lblNucleolus.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()

        Else
            lblNucleolus.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub





    Private Sub lblChloroplast_click(sender As Object, e As EventArgs) Handles lblChloroplast.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblChloroplast.Tag Then
            lblChloroplast.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()

        Else
            lblChloroplast.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub

    Private Sub lblChromosomes_click(sender As Object, e As EventArgs) Handles lblChromosomes.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblChromosomes.Tag Then
            lblChromosomes.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()

        Else
            lblChromosomes.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub

    Private Sub lblMitochrondrion_click(sender As Object, e As EventArgs) Handles lblMitochrondrion.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblMitochrondrion.Tag Then
            lblMitochrondrion.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()

        Else
            lblMitochrondrion.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub

    Private Sub lblCentriole_Click(sender As Object, e As EventArgs) Handles lblCentriole.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblCentriole.Tag Then
            lblCentriole.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()

        Else
            lblCentriole.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub

    Private Sub lblLysosome_click(sender As Object, e As EventArgs) Handles lblLysosome.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblLysosome.Tag Then
            lblLysosome.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()

        Else
            lblLysosome.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub

    Private Sub lblGolgiApparatus_click(sender As Object, e As EventArgs) Handles lblGolgiApparatus.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblGolgiApparatus.Tag Then
            lblGolgiApparatus.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()

        Else
            lblGolgiApparatus.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub

    Private Sub lblCellWall_click(sender As Object, e As EventArgs) Handles lblCellWall.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblCellWall.Tag Then
            lblCellWall.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()


        Else
            lblCellWall.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub

    Private Sub lblCellMembrane_Click(sender As Object, e As EventArgs) Handles lblCellMembrane.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblCellMembrane.Tag Then
            lblCellMembrane.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()


        Else
            lblCellMembrane.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub

    Private Sub btnCytoplasm_Click(sender As Object, e As EventArgs) Handles lblCytoplasm.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblCytoplasm.Tag Then
            lblCytoplasm.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()


        Else
            lblCytoplasm.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub

    Private Sub lblRibosome_Click(sender As Object, e As EventArgs) Handles lblRibosome.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblRibosome.Tag Then
            lblRibosome.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()


        Else
            lblRibosome.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub

    Private Sub lblEndoplasm_click(sender As Object, e As EventArgs) Handles lblEndoplasm.Click
        Dim intOrganelleCount2 As Integer
        If Label62.Tag = lblEndoplasm.Tag Then
            lblEndoplasm.Visible = False
            intOrganelleCorrectCounter = intOrganelleCorrectCounter + 1
            Do
                lblArrayOrganelle(intOrganelleCount2).Enabled = True
                intOrganelleCount2 = intOrganelleCount2 + 1
            Loop Until intOrganelleCount2 = 14
            Call QSetUp()


        Else
            lblEndoplasm.Enabled = False
            Call WrongOrganelle()
        End If
    End Sub

    Private intWrongOrg As Integer
    Private Sub WrongOrganelle()
        intWrongOrg = intWrongOrg + 1
        If intWrongOrg = 3 Then
            MsgBox("Good effort! You got a score of:" & Space(2) & intOrganelleCorrectCounter)

            Call OrganelleReset()
        End If
    End Sub

    Private Sub OrganelleReset()
        Dim IntCountReset As Integer

        tmrDNAGame.Enabled = False
        GroupBox1.Visible = True
        intOrganelleCorrectCounter = 0
        Timer10.Enabled = True
        Label62.Text = "Select the organelle that matches the given description."
        Label63.Visible = True
        lblNucleolus.Visible = False
        lblChromosomes.Visible = False
        lblChloroplast.Visible = False
        lblCentriole.Visible = False
        lblCellWall.Visible = False
        lblEndoplasm.Visible = False
        lblLysosome.Visible = False
        lblCellMembrane.Visible = False
        lblRibosome.Visible = False
        lblNucleolus.Visible = False
        lblMitochrondrion.Visible = False
        lblCytoplasm.Visible = False
        lblVacuole.Visible = False
        lblGolgiApparatus.Visible = False
        lblNucleus.Visible = False

        lblNucleolus.Enabled = True
        lblChromosomes.Enabled = True
        lblChloroplast.Enabled = True
        lblCentriole.Enabled = True
        lblCellWall.Enabled = True
        lblEndoplasm.Enabled = True
        lblLysosome.Enabled = True
        lblCellMembrane.Enabled = True
        lblRibosome.Enabled = True
        lblNucleolus.Enabled = True
        lblMitochrondrion.Enabled = True
        lblCytoplasm.Enabled = True
        lblVacuole.Enabled = True
        lblGolgiApparatus.Enabled = True
        lblNucleus.Enabled = True
        intWrongOrg = 0

        Do
            bolOrganelleArray(IntCountReset) = False
            IntCountReset = IntCountReset + 1
        Loop Until IntCountReset = 14



    End Sub

    Private Sub Label83_Click(sender As Object, e As EventArgs) Handles Label83.Click
        End
    End Sub

    
    Private Sub Label87_Click(sender As Object, e As EventArgs) Handles Label87.Click
        End
    End Sub

    Private Sub Label84_Click(sender As Object, e As EventArgs) Handles Label84.Click
        End
    End Sub

    Private Sub Label85_Click(sender As Object, e As EventArgs) Handles Label85.Click
        End
    End Sub

    Private bolSize9, bolSize10, bolSize11 As Boolean
    Private inExtra9, intExtra10, intExtra11 As Integer
    Private Sub tmrBiologyMainFont_Tick(sender As Object, e As EventArgs) Handles tmrBiologyMainFont.Tick
        If bolSize9 = False Then
            inExtra9 = inExtra9 + 1

            If inExtra9 > 10 Then
                bolSize9 = True
            End If
        Else
            inExtra9 = inExtra9 - 1
            If inExtra9 < 0 Then
                bolSize9 = False
            End If
        End If

        lblBioMain.Font = New Font("Kristen ITC", 72 + inExtra9)
    End Sub

    Private Sub tmrMathMainFont_Tick(sender As Object, e As EventArgs) Handles tmrMathMainFont.Tick
        If bolSize10 = False Then
            intExtra10 = intExtra10 + 1

            If intExtra10 > 10 Then
                bolSize10 = True
            End If
        Else
            intExtra10 = intExtra10 - 1
            If intExtra10 < 0 Then
                bolSize10 = False
            End If
        End If

        lblMathMainMenu.Font = New Font("Kristen ITC", 72 + intExtra10)
    End Sub

    Private Sub TabMain1_Click(sender As Object, e As EventArgs) Handles TabMain1.Click

    End Sub

    Private Sub tmrPhysicsMainFont_Tick(sender As Object, e As EventArgs) Handles tmrPhysicsMainFont.Tick
        If bolSize11 = False Then
            intExtra11 = intExtra11 + 1

            If intExtra11 > 10 Then
                bolSize11 = True
            End If
        Else
            intExtra11 = intExtra11 - 1
            If intExtra11 < 0 Then
                bolSize11 = False
            End If
        End If

        lblPhysicsFlash.Font = New Font("Kristen ITC", 72 + intExtra11)
    End Sub





    'Trigonometry Begins Here
    Private Sub Label76_Click(sender As Object, e As EventArgs) Handles Label76.Click
        Timer11.Enabled = True

        LineShape1.Visible = False
        LineShape2.Visible = False
        LineShape3.Visible = False
        LineShape4.Visible = False
        LineShape5.Visible = False
        LineShape6.Visible = False
        LineShape7.Visible = False
        Label94.Visible = False
        Label93.Visible = False
        Label92.Visible = True

        LineShape8.Visible = False
        LineShape9.Visible = False
        LineShape10.Visible = False
        LineShape11.Visible = False
        LineShape12.Visible = False
        LineShape13.Visible = False
        LineShape14.Visible = False
        LineShape15.Visible = False
        LineShape16.Visible = False
        LineShape17.Visible = False
        LineShape18.Visible = False
        LineShape19.Visible = False
        LineShape20.Visible = False
        LineShape21.Visible = False
        LineShape22.Visible = False
        LineShape23.Visible = False
        TextBox1.Visible = False
        Label95.Visible = False
        Button2.Visible = False

        Label89.Visible = False
        Label90.Visible = False
        Label91.Visible = False
        Label96.Visible = False
        Label97.Visible = False
        Label98.Visible = False
        Label99.Visible = False
        Label100.Visible = False
        Button2.Enabled = True
        TextBox1.Text = ""

        TabsMain.SelectTab(AAA)
        ' Reset Trig


    End Sub

    Private intAnswer As Integer
    Private Sub Trig()
        LineShape1.Visible = True
        LineShape2.Visible = True
        LineShape3.Visible = True

        Me.Refresh()

        Do
            LineShape1.X1 = 682 * Rnd()
        Loop Until LineShape1.X1 > 77
        Do
            LineShape1.Y1 = 772 * Rnd()
        Loop Until LineShape1.Y1 > 342
        Do
            LineShape1.X2 = 682 * Rnd()
        Loop Until LineShape1.X2 > 77 And (LineShape1.X2 - LineShape1.X1 > 150 Or LineShape1.X2 - LineShape1.X1 < -150)
        Do
            LineShape1.Y2 = 772 * Rnd()
        Loop Until LineShape1.Y2 > 342 And (LineShape1.Y2 - LineShape1.Y1 > 150 Or LineShape1.Y2 - LineShape1.Y1 < -150)

        Dim VertexX, VertexY As Integer
        VertexX = LineShape1.X1
        VertexY = LineShape1.Y2

        LineShape2.X1 = LineShape1.X1
        LineShape2.Y1 = LineShape1.Y1
        LineShape2.X2 = VertexX
        LineShape2.Y2 = VertexY

        LineShape3.X1 = LineShape1.X2
        LineShape3.Y1 = LineShape1.Y2
        LineShape3.X2 = VertexX
        LineShape3.Y2 = VertexY

        Dim intMidLine2 As Integer
        intMidLine2 = (LineShape2.Y1 + LineShape2.Y2) / 2
        If LineShape2.X1 > LineShape1.X2 Then
            Label89.Location = New Point(LineShape2.X1 + 20, intMidLine2)
        Else
            Label89.Location = New Point(LineShape2.X1 - 50, intMidLine2)
        End If
        If LineShape2.Y1 > LineShape2.Y2 Then
            Label89.Text = LineShape2.Y1 - LineShape2.Y2
        Else
            Label89.Text = LineShape2.Y2 - LineShape2.Y1
        End If

        Dim intMidLine3
        intMidLine3 = (LineShape3.X1 + LineShape3.X2) / 2 'This works i thnk
        If LineShape3.Y1 < LineShape2.Y1 Then
            Label90.Location = New Point(intMidLine3, LineShape3.Y1 - 30)
        Else
            Label90.Location = New Point(intMidLine3, LineShape3.Y1 + 20)
        End If
        If LineShape3.X1 > LineShape3.X2 Then
            Label90.Text = LineShape3.X1 - LineShape3.X2
        Else
            Label90.Text = LineShape3.X2 - LineShape3.X1
        End If
        intAnswer = Math.Sqrt(Label89.Text ^ 2 + Label90.Text ^ 2)
        Label91.Text = intAnswer

        Dim intMidLine1X, intMidLine1Y As Integer
        intMidLine1X = (LineShape1.X1 + LineShape1.X2) / 2
        intMidLine1Y = (LineShape1.Y1 + LineShape1.Y2) / 2

        If intMidLine1Y > VertexY Then
            intMidLine1Y = intMidLine1Y + 35
        Else
            intMidLine1Y = intMidLine1Y - 35
        End If
        If intMidLine1X > VertexX Then
            intMidLine1X = intMidLine1X + 35
        Else
            intMidLine1X = intMidLine1X - 35
        End If
        If intMidLine1X > VertexX And intMidLine1Y > VertexY Then
            intMidLine1X = intMidLine1X - 10
            intMidLine1Y = intMidLine1Y - 10
        End If

        Label91.Location = New Point(intMidLine1X, intMidLine1Y)

        Dim intRandom As Integer

        intRandom = 2 * Rnd()
        Select Case intRandom
            Case 0
                Label89.Visible = False
                Label90.Visible = True
                Label91.Visible = True
                intAnswer = Label89.Text
            Case 1
                Label89.Visible = True
                Label90.Visible = False
                Label91.Visible = True
                intAnswer = Label90.Text
            Case 2
                Label89.Visible = True
                Label90.Visible = True
                Label91.Visible = False
                intAnswer = Label91.Text
        End Select
    End Sub


    Private Sub Label92_Click(sender As Object, e As EventArgs) Handles Label92.Click
        Label92.Visible = False
        Call Refreshers()
    End Sub

    Private Sub Refreshers()
        LineShape1.Visible = True
        LineShape2.Visible = True
        LineShape3.Visible = True
        LineShape4.Visible = True
        LineShape5.Visible = True
        LineShape6.Visible = True
        LineShape7.Visible = True
        Label94.Visible = True
        Label93.Visible = True

        LineShape8.Visible = True
        LineShape9.Visible = True
        LineShape10.Visible = True
        LineShape11.Visible = True
        LineShape12.Visible = True
        LineShape13.Visible = True
        LineShape14.Visible = True
        LineShape15.Visible = True
        LineShape16.Visible = True
        LineShape17.Visible = True
        LineShape18.Visible = True
        LineShape19.Visible = True
        LineShape20.Visible = True
        LineShape21.Visible = True
        LineShape22.Visible = True
        LineShape23.Visible = True
        TextBox1.Visible = True
        Label95.Visible = True
        Button2.Visible = True

        Me.Refresh() 'use if stuff gets weird, in this version, its weird
        Call Trig()
    End Sub

    Private Sub Label94_Click(sender As Object, e As EventArgs) Handles Label94.Click
        Call Refreshers()
    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim intGuess As Integer
        If TextBox1.Text = "" Then
            intGuess = 1
        Else
            intGuess = TextBox1.Text
        End If
        If intGuess = intAnswer Or intGuess = intAnswer - 1 Or intGuess = intAnswer + 1 Then
            LineShape1.Visible = False
            LineShape1.Visible = False
            LineShape2.Visible = False
            LineShape2.Visible = False
            LineShape3.Visible = False
            LineShape3.Visible = False

            MsgBox("Congrats")
            Label96.Text = ""
            Label97.Text = ""
            Label98.Text = ""
            Label99.Text = ""
            Call Trig()
        Else
            MsgBox("Wrong, Here's the Solution")
            Label96.Visible = True
            Label97.Visible = True
            Label98.Visible = True
            Label99.Visible = True
            If Label89.Visible = True Then
                Label96.Text = "Side A = " & Label89.Text
            Else
                Label96.Text = "Side A = Missing"
                Label99.Text = "Missing Side = " & Label89.Text
            End If
            If Label90.Visible = True Then
                Label97.Text = "Side B = " & Label90.Text
            Else
                Label97.Text = "Side B = Missing"
                Label99.Text = "Missing Side = " & Label90.Text
            End If
            If Label91.Visible = True Then
                Label98.Text = "Side C = " & Label91.Text
            Else
                Label98.Text = "Side C = Missing"
                Label99.Text = "Missing Side = " & Label91.Text
            End If
            Timer12.Enabled = True
            Label100.Visible = True
            Button2.Enabled = False

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If IsNumeric(TextBox1.Text) = False Then
            TextBox1.Text = ""
        ElseIf TextBox1.Text / 1000 >= 1 Then
            TextBox1.Text = ""
        End If
    End Sub

    Dim bolSize12 As Boolean
    Dim intExtra12 As Integer
    Private Sub Timer7_Tick(sender As Object, e As EventArgs) Handles Timer7.Tick
        If bolSize12 = False Then
            intExtra12 = intExtra12 + 1

            If intExtra12 > 10 Then
                bolSize12 = True
            End If
        Else
            intExtra12 = intExtra12 - 1
            If intExtra12 < 0 Then
                bolSize12 = False
            End If
        End If

        Label49.Font = New Font("Kristen ITC", 36 + intExtra12)
    End Sub

    Dim bolSize13 As Boolean
    Dim intExtra13 As Integer
    Private Sub Timer11_Tick(sender As Object, e As EventArgs) Handles Timer11.Tick
        If bolSize13 = False Then
            intExtra13 = intExtra13 + 1

            If intExtra13 > 10 Then
                bolSize13 = True
            End If
        Else
            intExtra13 = intExtra13 - 1
            If intExtra13 < 0 Then
                bolSize13 = False
            End If
        End If

        Label92.Font = New Font("Kristen ITC", 36 + intExtra13)
    End Sub

    Dim bolSize14 As Boolean
    Dim intExtra14 As Integer
    Private Sub Timer12_Tick(sender As Object, e As EventArgs) Handles Timer12.Tick
        If bolSize14 = False Then
            intExtra14 = intExtra14 + 1

            If intExtra14 > 4 Then
                bolSize14 = True
            End If
        Else
            intExtra14 = intExtra14 - 1
            If intExtra14 < 0 Then
                bolSize14 = False
            End If
        End If

        Label100.Font = New Font("Kristen ITC", 16 + intExtra14)
    End Sub

    Private Sub Label100_Click(sender As Object, e As EventArgs) Handles Label100.Click
        Label100.Visible = False
        Label96.Visible = False
        Label97.Visible = False
        Label98.Visible = False
        Label99.Visible = False
        Button2.Enabled = True

        Call Trig()
    End Sub
End Class

