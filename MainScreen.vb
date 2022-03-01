Imports System.Windows.Forms
Imports Marshall_Mix_Design__M.B.M.K_.Analaysis
Imports Marshall_Mix_Design__M.B.M.K_.CriteriaForm
Imports Marshall_Mix_Design__M.B.M.K_.DataShow
Imports Marshall_Mix_Design__M.B.M.K_.Degree_Pol_text
Imports Marshall_Mix_Design__M.B.M.K_.Form1
Imports Marshall_Mix_Design__M.B.M.K_.Form2
Imports System.Threading
Public Class Degree_Pol_text
    Public save As Integer
    Public AA, BB, CC, DD, EE, FF, draw_scale As Double

    Dim fontt As Font = New Font(New FontFamily("Arial"), 12, FontStyle.Bold)
    REM ///////// Text Import Data Variables //////////////
    Dim Subcount As Integer    REM Cout for separate INPUT e.g AC,A,B,C,S,F "FROM 1 TO 6"
    Dim No_Sample As Integer   REM Number of sample
    Dim current_field As Double REM As indicator for empty field e.g ,,,, to get zero to know new No.sample :)
    Dim coun_No_of_TRY(100) As Integer REM to count number of trails for each asphalt content to find an    AVG e.g (1+2+2)/coun_No_of_TRY(100)=3
    Dim No_of_TRY As Integer   REM Number of try e.g sample(1) with 3% AC have 3 trails 1 , 2 , 3 then GET AVG
    Dim Ac11(100, 100) As Double REM AC = Asphalt Contant in as a percantage e.g 3%
    Dim a11(100, 100) As Double  REM A = Dry Weight of sample in GRAMS
    Dim b11(100, 100) As Double  REM B = SSD Weight of sample in GRAMS
    Dim c11(100, 100) As Double  REM C = In Water Weight of sample in GRAMS
    Dim S(100, 100) As Double  REM S = Corrected Stability 
    Dim F(100, 100) As Double  REM F = Flow
    REM ///////// Figures and Analyze Data Variables //////////////
    Dim AC_Analyze(100) As Double
    Dim Gmb_Analyze(100) As Double
    Dim Gmm_Analyze(100) As Double
    Dim AV_Percentage_Analyze(100) As Double
    Dim VMA_Percentage_Analyze(100) As Double
    Dim S_Analyze(100) As Double
    Dim F_Analyze(100) As Double
    REM ///////// Criteria Variables //////////////
    Dim traffic_volum As String
    Public Min_stability As Double
    Public flow_lower_limit As Double
    Public flow_upper_limit As Double
    Public AV_Percentage_lower_percentage As Double
    Public AV_Percentage_Upper_percentage As Double
    Public Min_VMA_Percentage As Double
    REM /////////////////////Curve Fitting Variables//////////////////////////////
    Dim x(1000) As Double
    Dim y(1000) As Double
    Dim s1, s2 As Double
    Dim fv As Double
    Dim sumy As Double
    Dim st, sr As Double
    Dim yavg, maxx, minx, maxy, miny As Double
    Dim number_datta, number_fit As Double
    Dim iii As Double REM represent No_Sample
    Dim n, rr, nn, ccc, M, av, bv As Double REM represent Matrix Size
    Dim A_Gma(1000), A_S(1000), A_F(1000), A_AV(1000), A_VMA(1000) As Double  REM it is a a0 , a1 , a2 for curve fitting for each properties
    Dim minx_Gma, minx_S, minx_F, minx_AV, minx_VMA As Double
    Dim maxx_Gma, maxx_S, maxx_F, maxx_AV, maxx_VMA As Double
    Dim miny_Gma, miny_S, miny_F, miny_AV, miny_VMA As Double
    Dim maxy_Gma, maxy_S, maxy_F, maxy_AV, maxy_VMA As Double
    Dim degree_ploy_draw_curve As Integer
    Dim gmb_deg_ploy_txt, s_deg_plo_txt, f_deg_plo_txt, av_deg_plo_txt, vma_deg_plo_txt As Integer

    Dim x_draw, y_draw As Integer
    Dim x_scale, y_scale As Double
    Dim Max_X_valu_draw, Min_X_valu_draw, Max_Y_valu_draw, Min_Y_valu_draw As Double

    REM ////////////////////// Analysis variables///////////////////////////////////////
    Dim optimum_Asphalt_content, op1, op2, op3, dddd As Double
    REM /////////////////////////////Result Variables////////////////////////////////////
    Dim OCA As Double
    Dim gma_result As Double
    Dim s_result As Double
    Dim f_result As Double
    Dim av_result As Double
    Dim vma_result As Double
    Dim Max_Gmb
    REM /////////////////////////////////////////////////////////////////////////////////

    Private Sub MMD_Load11(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        REM when Software first load what do ? ? ? 
        gmb_deg_ploy_txt = 1 : s_deg_plo_txt = 1 : f_deg_plo_txt = 1 : av_deg_plo_txt = 1 : vma_deg_plo_txt = 1
        draw_scale = 1
        gmb_deg_plo.Text = gmb_deg_ploy_txt
        s_deg_plo.Text = s_deg_plo_txt
        f_deg_plo.Text = f_deg_plo_txt
        av_deg_plo.Text = av_deg_plo_txt
        vma_deg_plo.Text = vma_deg_plo_txt
        Label9.Hide() : Label10.Hide() : Label11.Hide() : Label12.Hide() : Label13.Hide() : GroupBox1.Hide()
        gmb_deg_plo.Hide() : gmb_dec.Hide() : gmb_inc.Hide()
        s_deg_plo.Hide() : s_dec.Hide() : s_inc.Hide()
        f_deg_plo.Hide() : f_dec.Hide() : f_inc.Hide()
        av_deg_plo.Hide() : av_dec.Hide() : av_inc.Hide()
        vma_deg_plo.Hide() : vma_dec.Hide() : vma_inc.Hide()

        Button_Run.Enabled = False
        Button_CriteriaModified.Enabled = False

        gb_txt.Text = "1.02"
        gsb_txt.Text = "2.724"
        gse_txt.Text = "2.777"
        Me.WindowState = FormWindowState.Maximized
        CheckBox_Medium.Checked = True
        CheckBox_N.Checked = True
        CheckBox_Unit.Checked = True
        Button_Data_Analysis.Enabled = False
        Button_ShowData.Enabled = False

        My.Forms.CriteriaForm.Show()
        My.Forms.CriteriaForm.Close() REM to get information open then process then get result!1



    End Sub
    Private Sub Import_Button_Click(sender As System.Object, e As System.EventArgs) Handles Import_Button.Click
        OpenFileDialog1.ShowDialog()
        No_Sample = 1
        No_of_TRY = 1
        Subcount = 1
        Try
            Dim reader As New IO.StreamReader(OpenFileDialog1.FileName)
            Using MyReader As New Microsoft.VisualBasic.
                            FileIO.TextFieldParser(
                              reader)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim currentRow As String()

                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()
                        Dim currentField As Double
                        For Each currentField In currentRow
                            If (Subcount = 1 And currentField <> 0) Then
                                Ac11(No_Sample, No_of_TRY) = currentField
                            End If
                            If (Subcount = 2 And currentField <> 0) Then
                                a11(No_Sample, No_of_TRY) = currentField
                            End If
                            If (Subcount = 3 And currentField <> 0) Then
                                b11(No_Sample, No_of_TRY) = currentField
                            End If
                            If (Subcount = 4 And currentField <> 0) Then
                                c11(No_Sample, No_of_TRY) = currentField
                            End If
                            If (Subcount = 5 And currentField <> 0) Then
                                S(No_Sample, No_of_TRY) = currentField
                            End If
                            If (Subcount = 6 And currentField <> 0) Then
                                F(No_Sample, No_of_TRY) = currentField
                            End If
                            Subcount = Subcount + 1
                            current_field = currentField
                        Next
                    Catch ex As Microsoft.VisualBasic.
                                FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
                        "is not valid and will be skipped.")
                    End Try
                    If (current_field = 0) Then
                        coun_No_of_TRY(No_Sample) = No_of_TRY - 1
                        No_Sample = No_Sample + 1
                        No_of_TRY = 0
                    End If
                    Subcount = 1
                    No_of_TRY = No_of_TRY + 1

                End While
            End Using
            Button_ShowData.Enabled = True
            Button_CriteriaModified.Enabled = True
        Catch ex As Exception
            MsgBox("Check Your Data Input!")
        End Try
        No_Sample -= 1
    End Sub
    Private Sub Show_Data_btn_Click(sender As System.Object, e As System.EventArgs) Handles Button_ShowData.Click
        Dim k As Integer REM to un_repeated of No.Sample
        For i = 1 To No_Sample
            k = i
            For j = 1 To coun_No_of_TRY(i)
                If (k = i) Then
                    DataShow.DataGridView1.Rows.Add(i, j, Ac11(i, j), a11(i, j), b11(i, j), c11(i, j), S(i, j), F(i, j))
                    k = k + 1
                Else
                    DataShow.DataGridView1.Rows.Add("", j, Ac11(i, j), a11(i, j), b11(i, j), c11(i, j), S(i, j), F(i, j))
                End If

            Next
        Next
        DataShow.WindowState = FormWindowState.Normal
        DataShow.Show()
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button_CriteriaModified.Click
        CriteriaForm.Show()
    End Sub
    Private Sub CheckBox_Heavy_CheckedChanged11(sender As System.Object, e As System.EventArgs) Handles CheckBox_Heavy.CheckedChanged
        If (CheckBox_Heavy.Checked = True) Then
            CheckBox_Medium.Checked = False
            CheckBox_Light.Checked = False
        End If
    End Sub
    Private Sub CheckBox_Medium_CheckedChanged11(sender As System.Object, e As System.EventArgs) Handles CheckBox_Medium.CheckedChanged
        If (CheckBox_Medium.Checked = True) Then
            CheckBox_Heavy.Checked = False
            CheckBox_Light.Checked = False
        End If
    End Sub
    Private Sub CheckBox_Light_CheckedChanged11(sender As System.Object, e As System.EventArgs) Handles CheckBox_Light.CheckedChanged
        If (CheckBox_Light.Checked = True) Then
            CheckBox_Medium.Checked = False
            CheckBox_Heavy.Checked = False
        End If
    End Sub
    Private Sub CheckBox_N_CheckedChanged11(sender As System.Object, e As System.EventArgs) Handles CheckBox_N.CheckedChanged
        If (CheckBox_N.Checked = True) Then
            CheckBox_Kg.Checked = False
            CheckBox_Ib.Checked = False
        End If
    End Sub
    Private Sub CheckBox_Kg_CheckedChanged11(sender As System.Object, e As System.EventArgs) Handles CheckBox_Kg.CheckedChanged
        If (CheckBox_Kg.Checked = True) Then
            CheckBox_N.Checked = False
            CheckBox_Ib.Checked = False
        End If
    End Sub
    Private Sub CheckBox_Ib_CheckedChanged11(sender As System.Object, e As System.EventArgs) Handles CheckBox_Ib.CheckedChanged
        If (CheckBox_Ib.Checked = True) Then
            CheckBox_Kg.Checked = False
            CheckBox_N.Checked = False
        End If
    End Sub
    Private Sub CheckBox_Unit_CheckedChanged11(sender As System.Object, e As System.EventArgs) Handles CheckBox_Unit.CheckedChanged
        If (CheckBox_Unit.Checked = True) Then
            CheckBox_25mm.Checked = False
            CheckBox_001in.Checked = False
        End If
    End Sub
    Private Sub CheckBox_25mm_CheckedChanged11(sender As System.Object, e As System.EventArgs) Handles CheckBox_25mm.CheckedChanged
        If (CheckBox_25mm.Checked = True) Then
            CheckBox_Unit.Checked = False
            CheckBox_001in.Checked = False
        End If
    End Sub
    Private Sub CheckBox_001in_CheckedChanged11(sender As System.Object, e As System.EventArgs) Handles CheckBox_001in.CheckedChanged
        If (CheckBox_001in.Checked = True) Then
            CheckBox_25mm.Checked = False
            CheckBox_Unit.Checked = False
        End If
    End Sub
    Private Sub Button_Run_Click(sender As System.Object, e As System.EventArgs) Handles Button_Run.Click
        REM *******Data Analyaz Calculation**************
        REM 1) Asphalt content Groups *************************
        Me.CreateGraphics.Clear(BackColor)
        clear_result()
        Label9.Show() : Label10.Show() : Label11.Show() : Label12.Show() : Label13.Show() : GroupBox1.Show()
        gmb_deg_plo.Show() : gmb_dec.Show() : gmb_inc.Show()
        s_deg_plo.Show() : s_dec.Show() : s_inc.Show()
        f_deg_plo.Show() : f_dec.Show() : f_inc.Show()
        av_deg_plo.Show() : av_dec.Show() : av_inc.Show()
        vma_deg_plo.Show() : vma_dec.Show() : vma_inc.Show()
        Try
            Dim val As Double
            val = 0
            For i = 1 To No_Sample
                For j = 1 To coun_No_of_TRY(i)
                    val = val + (Ac11(i, j))
                Next
                AC_Analyze(i) = val / coun_No_of_TRY(i)
                val = 0
            Next
            REM 2) *********** Gmb Groups *************************
            val = 0
            For i = 1 To No_Sample
                For j = 1 To coun_No_of_TRY(i)
                    val = val + (a11(i, j) / (b11(i, j) - c11(i, j)))
                Next
                Gmb_Analyze(i) = val / coun_No_of_TRY(i)
                val = 0
            Next
            REM 3) *********** S Groups *************************
            val = 0
            For i = 1 To No_Sample
                For j = 1 To coun_No_of_TRY(i)
                    val = val + S(i, j)
                Next
                S_Analyze(i) = val / coun_No_of_TRY(i)
                val = 0
            Next
            REM 4) *********** Flow Groups ***********************
            val = 0
            For i = 1 To No_Sample
                For j = 1 To coun_No_of_TRY(i)
                    val = val + F(i, j)
                Next
                F_Analyze(i) = val / coun_No_of_TRY(i)
                val = 0
            Next
            REM 5) *********** Gmm Groups ***********************
            For i = 1 To No_Sample
                Gmm_Analyze(i) = 100 / (((100 - AC_Analyze(i)) / gse_txt.Text) + (AC_Analyze(i) / gb_txt.Text))
                AV_Percentage_Analyze(i) = (1 - (Gmb_Analyze(i) / Gmm_Analyze(i))) * 100
                VMA_Percentage_Analyze(i) = (1 - (1 - AC_Analyze(i) / 100) * (Gmb_Analyze(i) / gse_txt.Text)) * 100
            Next
            Button_Data_Analysis.Enabled = True
        Catch ex As Exception
            MsgBox("Check Your Data!!")
        End Try


        fitt("gma", gmb_deg_ploy_txt)
        Draw_Fitting_Curve("gma", 275, 310)

        fitt("s", s_deg_plo_txt)
        Draw_Fitting_Curve("s", 675, 310)

        fitt("flow", f_deg_plo_txt)
        Draw_Fitting_Curve("flow", 1075, 310)

        fitt("av", av_deg_plo_txt)
        Draw_Fitting_Curve("av", 275, 670)

        fitt("vma", vma_deg_plo_txt)
        Draw_Fitting_Curve("vma", 675, 670)

        optimum_ac()
        TextBox_rslt_oca.Text = Format(optimum_Asphalt_content, "0.00")


    End Sub
    Private Sub Button_Data_Analysis_Click(sender As System.Object, e As System.EventArgs) Handles Button_Data_Analysis.Click
        For i = 1 To No_Sample
            Analaysis.DataGridView1.Rows.Add(i, Format(AC_Analyze(i), "0.00"), Format(Gmb_Analyze(i), "0.00"), Format(Gmm_Analyze(i), "0.00"), Format(AV_Percentage_Analyze(i), "0.00"), Format(S_Analyze(i), "0.00"), Format(F_Analyze(i), "0.00"), Format(VMA_Percentage_Analyze(i), "0.00"))
        Next
        Analaysis.Show()
    End Sub
    Sub fitt(key As String, Degree_Poly As Integer)
        If (Degree_Poly >= (No_Sample - 1)) Then Degree_Poly = No_Sample - 1
        degree_ploy_draw_curve = Degree_Poly
        If (key = "gma") Then
            For i = 0 To No_Sample - 1
                x(i) = AC_Analyze(i + 1) : y(i) = Gmb_Analyze(i + 1)
            Next
        End If
        If (key = "s") Then
            For i = 0 To No_Sample - 1
                x(i) = AC_Analyze(i + 1) : y(i) = S_Analyze(i + 1)
            Next
        End If
        If (key = "flow") Then
            For i = 0 To No_Sample - 1
                x(i) = AC_Analyze(i + 1) : y(i) = F_Analyze(i + 1)
            Next
        End If
        If (key = "av") Then
            For i = 0 To No_Sample - 1
                x(i) = AC_Analyze(i + 1) : y(i) = AV_Percentage_Analyze(i + 1)
            Next
        End If
        If (key = "vma") Then
            For i = 0 To No_Sample - 1
                x(i) = AC_Analyze(i + 1) : y(i) = VMA_Percentage_Analyze(i + 1)
            Next
        End If
        REM STEP ***1) //FIND MAX MIN FOR NUMBER AND Avg  
        maxx = 0
        maxy = 0
        minx = 10000000
        miny = 10000000
        iii = No_Sample
        number_datta = iii
        sumy = 0
        For ii = 0 To iii - 1
            sumy = sumy + y(ii)
            If x(ii) > maxx Then maxx = x(ii)
            If Math.Abs(y(ii)) > maxy Then maxy = Math.Abs(y(ii))
            If x(ii) < minx Then minx = x(ii)
            If y(ii) < miny Then miny = y(ii)
        Next
        yavg = sumy / number_datta
        st = 0
        For i = 0 To (number_datta - 1)
            st = st + (y(i) - yavg) ^ 2
        Next
        REM //////////////////////////////////////////////
        number_fit = degree_ploy_draw_curve
        n = number_fit + 1
        Dim sumx(n * n) As Double

        For i = 0 To (2 * number_fit)
            For j = 0 To (number_datta - 1)
                sumx(i) = sumx(i) + x(j) ^ i
            Next j
        Next i

        Dim sumxy(n) As Double
        For i = 0 To number_fit
            For j = 0 To (number_datta - 1)
                sumxy(i) = sumxy(i) + ((x(j)) ^ i) * y(j)
            Next j
        Next i
        REM ////////////////////////////////////////////
        rr = n : nn = 1
        Dim a11(n, n) As Double
        Dim b11(n, n) As Double
        Dim c11(rr, nn) As Double
        Dim d11(n, nn) As Double
        For i = 0 To (n - 1)
            For j = 0 To (n - 1)
                If i = j Then
                    b11(i, j) = 1.0
                Else
                    b11(i, j) = 0
                End If
            Next j
        Next i
        REM ***************************************
        For i = 0 To (rr - 1)
            For j = 0 To (nn - 1)
                c11(i, j) = sumxy(i)
            Next j
        Next i
        REM **************************************
        For i = 0 To (n - 1)
            ccc = i
            For j = 0 To (n - 1)

                a11(i, j) = sumx(ccc)
                ccc = ccc + 1
            Next j
        Next i
        REM ////////////////////////////////////////
        For i = 1 To (n - 1)
            For j = 0 To (i - 1)
                fv = a11(i, j) / a11(j, j)
                For Me.M = 0 To n - 1
                    a11(i, Me.M) = a11(i, Me.M) - fv * a11(j, Me.M)
                    b11(i, Me.M) = b11(i, Me.M) - fv * b11(j, Me.M)
                Next
            Next
        Next

        For i = (n - 2) To 0 Step -1
            For j = (n - 1) To (i + 1) Step -1
                fv = a11(i, j) / a11(j, j)
                For Me.M = (n - 1) To 0 Step -1
                    a11(i, Me.M) = a11(i, Me.M) - fv * a11(j, Me.M)
                    b11(i, Me.M) = b11(i, Me.M) - fv * b11(j, Me.M)
                Next
            Next
        Next
        REM ////////////////SOLVE//////////////////////
        For i = 0 To (n - 1)
            For j = 0 To (n - 1)
                b11(i, j) = b11(i, j) / a11(i, i)
            Next
        Next
        av = 1
        bv = 1
        For r33 = 0 To (rr - 1)
            For c33 = 0 To (nn - 1)
                For nnn = 0 To (n - 1)
                    d11(r33, c33) = d11(r33, c33) + (b11(r33, nnn) * c11(nnn, c33))
                Next nnn
            Next c33
        Next
        REM /////////////// Final point get a0 + a1x^1 + a2x^2+.......anx^n for each properties depends user input :)///////////
        If (key = "gma") Then
            For i = 0 To n - 1
                A_Gma(i) = d11(i, 0)
            Next
            minx_Gma = minx
            maxx_Gma = maxx
            miny_Gma = miny
            maxy_Gma = maxy
        End If
        If (key = "s") Then
            For i = 0 To n - 1
                A_S(i) = d11(i, 0)
            Next
            minx_S = minx
            maxx_S = maxx
            miny_S = miny
            maxy_S = maxy
        End If
        If (key = "flow") Then
            For i = 0 To n - 1
                A_F(i) = d11(i, 0)
            Next
            minx_F = minx
            maxx_F = maxx
            miny_F = miny
            maxy_F = maxy
        End If
        If (key = "av") Then
            For i = 0 To n - 1
                A_AV(i) = d11(i, 0)
            Next
            minx_AV = minx
            maxx_AV = maxx
            miny_AV = miny
            maxy_AV = maxy
        End If
        If (key = "vma") Then
            For i = 0 To n - 1
                A_VMA(i) = d11(i, 0)
            Next
            minx_VMA = minx
            maxx_VMA = maxx
            miny_VMA = miny
            maxy_VMA = maxy
        End If
        REM ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    End Sub
    Sub Draw_Fitting_Curve(key As String, x As Integer, y As Integer)

        REM ///////////////////////////////////////

        If (key = "gma") Then Max_X_valu_draw = maxx_Gma : Min_X_valu_draw = minx_Gma : Max_Y_valu_draw = maxy_Gma : Min_Y_valu_draw = miny_Gma
        If (key = "s") Then Max_X_valu_draw = maxx_S : Min_X_valu_draw = minx_S : Max_Y_valu_draw = maxy_S : Min_Y_valu_draw = miny_S
        If (key = "flow") Then Max_X_valu_draw = maxx_F : Min_X_valu_draw = minx_F : Max_Y_valu_draw = maxy_F : Min_Y_valu_draw = miny_F
        If (key = "av") Then Max_X_valu_draw = maxx_AV : Min_X_valu_draw = minx_AV : Max_Y_valu_draw = maxy_AV : Min_Y_valu_draw = miny_AV
        If (key = "vma") Then Max_X_valu_draw = maxx_VMA : Min_X_valu_draw = minx_VMA : Max_Y_valu_draw = maxy_VMA : Min_Y_valu_draw = miny_VMA

        x_scale = (250.0 / (Max_X_valu_draw - Min_X_valu_draw))
        y_scale = (250.0 / (Max_Y_valu_draw - Min_Y_valu_draw))
        REM //////////////////////////////////////
        Dim mmm As String
        Dim sdf As Integer
        For i = (x - 25) To (x + 275) Step 25
            Me.CreateGraphics.DrawLine(Pens.Black, i, y + 25, i, y - 275)
        Next
        For i = (x) To (x + 250) Step 50
            Me.CreateGraphics.DrawString(Format(((i - x) / (x_scale)) + Min_X_valu_draw, "0.00"), fontt, Brushes.Blue, i - 15, y + 25)
        Next
        For i = (y + 25) To (y - 275) Step -25
            Me.CreateGraphics.DrawLine(Pens.Black, x - 25, i, x + 275, i)
        Next
        For i = (y) To (y - 250) Step -25
            mmm = Format((-(i - y) / (y_scale)) + Min_Y_valu_draw, "0.00")
            sdf = mmm.Length()

            Me.CreateGraphics.DrawString(Format((-(i - y) / (y_scale)) + Min_Y_valu_draw, "0.00"), fontt, Brushes.Blue, x - (sdf + 2) * 10, i - 10)

        Next
        REM /////////////////////////////////////
        Dim yy As Double
        If (key = "gma") Then
            For i = 1 To No_Sample
                x_draw = x + ((AC_Analyze(i) - Min_X_valu_draw) * x_scale)
                y_draw = y - ((Gmb_Analyze(i) - Min_Y_valu_draw) * y_scale)
                Me.CreateGraphics.FillEllipse(Brushes.Red, x_draw - 5, y_draw - 5, 10, 10)
            Next
            For i = AC_Analyze(1) To AC_Analyze(No_Sample) Step 0.005
                yy = 0
                For j = 0 To degree_ploy_draw_curve
                    yy = yy + (A_Gma(j) * (i) ^ j)
                Next
                x_draw = x + ((i - Min_X_valu_draw) * x_scale)
                y_draw = y - ((yy - Min_Y_valu_draw) * y_scale)
                Me.CreateGraphics.FillEllipse(Brushes.Black, x_draw - 2, y_draw - 2, 4, 4)
            Next
        End If
        If (key = "s") Then
            For i = 1 To No_Sample
                x_draw = x + ((AC_Analyze(i) - Min_X_valu_draw) * x_scale)
                y_draw = y - ((S_Analyze(i) - Min_Y_valu_draw) * y_scale)
                Me.CreateGraphics.FillEllipse(Brushes.Red, x_draw - 5, y_draw - 5, 10, 10)
            Next
            For i = AC_Analyze(1) To AC_Analyze(No_Sample) Step 0.005
                yy = 0
                For j = 0 To degree_ploy_draw_curve
                    yy = yy + (A_S(j) * (i) ^ j)
                Next
                x_draw = x + ((i - Min_X_valu_draw) * x_scale)
                y_draw = y - ((yy - Min_Y_valu_draw) * y_scale)
                Me.CreateGraphics.FillEllipse(Brushes.Black, x_draw - 2, y_draw - 2, 4, 4)
            Next
        End If
        If (key = "flow") Then
            For i = 1 To No_Sample
                x_draw = x + ((AC_Analyze(i) - Min_X_valu_draw) * x_scale)
                y_draw = y - ((F_Analyze(i) - Min_Y_valu_draw) * y_scale)
                Me.CreateGraphics.FillEllipse(Brushes.Red, x_draw - 5, y_draw - 5, 10, 10)
            Next
            For i = AC_Analyze(1) To AC_Analyze(No_Sample) Step 0.005
                yy = 0
                For j = 0 To degree_ploy_draw_curve
                    yy = yy + (A_F(j) * (i) ^ j)
                Next
                x_draw = x + ((i - Min_X_valu_draw) * x_scale)
                y_draw = y - ((yy - Min_Y_valu_draw) * y_scale)
                Me.CreateGraphics.FillEllipse(Brushes.Black, x_draw - 2, y_draw - 2, 4, 4)
            Next
        End If
        If (key = "av") Then
            For i = 1 To No_Sample
                x_draw = x + ((AC_Analyze(i) - Min_X_valu_draw) * x_scale)
                y_draw = y - ((AV_Percentage_Analyze(i) - Min_Y_valu_draw) * y_scale)
                Me.CreateGraphics.FillEllipse(Brushes.Red, x_draw - 5, y_draw - 5, 10, 10)
            Next
            For i = AC_Analyze(1) To AC_Analyze(No_Sample) Step 0.005
                yy = 0
                For j = 0 To degree_ploy_draw_curve
                    yy = yy + (A_AV(j) * (i) ^ j)
                Next
                x_draw = x + ((i - Min_X_valu_draw) * x_scale)
                y_draw = y - ((yy - Min_Y_valu_draw) * y_scale)
                Me.CreateGraphics.FillEllipse(Brushes.Black, x_draw - 2, y_draw - 2, 4, 4)
            Next
        End If
        If (key = "vma") Then
            For i = 1 To No_Sample
                x_draw = x + ((AC_Analyze(i) - Min_X_valu_draw) * x_scale)
                y_draw = y - ((VMA_Percentage_Analyze(i) - Min_Y_valu_draw) * y_scale)
                Me.CreateGraphics.FillEllipse(Brushes.Red, x_draw - 5, y_draw - 5, 10, 10)
            Next
            For i = AC_Analyze(1) To AC_Analyze(No_Sample) Step 0.005
                yy = 0
                For j = 0 To degree_ploy_draw_curve
                    yy = yy + (A_VMA(j) * (i) ^ j)
                Next
                x_draw = x + ((i - Min_X_valu_draw) * x_scale)
                y_draw = y - ((yy - Min_Y_valu_draw) * y_scale)
                Me.CreateGraphics.FillEllipse(Brushes.Black, x_draw - 2, y_draw - 2, 4, 4)
            Next
        End If
        REM ///////////////////////////////////
    End Sub

    Sub degree_change(key As String, s As Integer)
        Me.CreateGraphics.Clear(BackColor)
        If (key = "gma") Then
            gmb_deg_ploy_txt = gmb_deg_ploy_txt + s
            gmb_deg_plo.Text = gmb_deg_ploy_txt
        End If
        If (key = "s") Then
            s_deg_plo_txt = s_deg_plo_txt + s
            s_deg_plo.Text = s_deg_plo_txt
        End If
        If (key = "f") Then
            f_deg_plo_txt = f_deg_plo_txt + s
            f_deg_plo.Text = f_deg_plo_txt
        End If
        If (key = "av") Then
            av_deg_plo_txt = av_deg_plo_txt + s
            av_deg_plo.Text = av_deg_plo_txt
        End If
        If (key = "vma") Then
            vma_deg_plo_txt = vma_deg_plo_txt + s
            vma_deg_plo.Text = vma_deg_plo_txt
        End If
        fitt("gma", gmb_deg_ploy_txt)
        Draw_Fitting_Curve("gma", 275, 310)

        fitt("s", s_deg_plo_txt)
        Draw_Fitting_Curve("s", 675, 310)

        fitt("flow", f_deg_plo_txt)
        Draw_Fitting_Curve("flow", 1075, 310)

        fitt("av", av_deg_plo_txt)
        Draw_Fitting_Curve("av", 275, 670)

        fitt("vma", vma_deg_plo_txt)
        Draw_Fitting_Curve("vma", 675, 670)

    End Sub

    Private Sub gmb_inc_Click(sender As System.Object, e As System.EventArgs) Handles gmb_inc.Click
        degree_change("gma", 1)
        optimum_ac()
        TextBox_rslt_oca.Text = Format(optimum_Asphalt_content, "0.00")
        clear_result()
    End Sub

    Private Sub gmb_dec_Click(sender As System.Object, e As System.EventArgs) Handles gmb_dec.Click
        degree_change("gma", -1)
        optimum_ac()
        clear_result()
    End Sub

    Private Sub s_inc_Click(sender As System.Object, e As System.EventArgs) Handles s_inc.Click
        degree_change("s", 1)
        optimum_ac()
        TextBox_rslt_oca.Text = Format(optimum_Asphalt_content, "0.00")
        clear_result()
    End Sub

    Private Sub s_dec_Click(sender As System.Object, e As System.EventArgs) Handles s_dec.Click
        degree_change("s", -1)
        optimum_ac()
        TextBox_rslt_oca.Text = Format(optimum_Asphalt_content, "0.00")
        clear_result()
    End Sub

    Private Sub f_inc_Click(sender As System.Object, e As System.EventArgs) Handles f_inc.Click
        degree_change("f", 1)
        optimum_ac()
        TextBox_rslt_oca.Text = Format(optimum_Asphalt_content, "0.00")
        clear_result()
    End Sub

    Private Sub f_dec_Click(sender As System.Object, e As System.EventArgs) Handles f_dec.Click
        degree_change("f", -1)
        optimum_ac()
        TextBox_rslt_oca.Text = Format(optimum_Asphalt_content, "0.00")
        clear_result()
    End Sub

    Private Sub av_inc_Click(sender As System.Object, e As System.EventArgs) Handles av_inc.Click
        degree_change("av", 1)
        optimum_ac()
        TextBox_rslt_oca.Text = Format(optimum_Asphalt_content, "0.00")
        clear_result()
    End Sub

    Private Sub av_dec_Click(sender As System.Object, e As System.EventArgs) Handles av_dec.Click
        degree_change("av", -1)
        optimum_ac()
        TextBox_rslt_oca.Text = Format(optimum_Asphalt_content, "0.00")
        clear_result()
    End Sub

    Private Sub vma_inc_Click(sender As System.Object, e As System.EventArgs) Handles vma_inc.Click
        degree_change("vma", 1)
        optimum_ac()
        TextBox_rslt_oca.Text = Format(optimum_Asphalt_content, "0.00")
        clear_result()
    End Sub

    Private Sub vma_dec_Click(sender As System.Object, e As System.EventArgs) Handles vma_dec.Click
        degree_change("vma", -1)
        optimum_ac()
        TextBox_rslt_oca.Text = Format(optimum_Asphalt_content, "0.00")
        clear_result()
    End Sub

    Sub optimum_ac()
        Dim maxim As Double
        Dim optimum_gma, optimum_s, optimum_av, optimum_all As Double
        REM ///////////////optimum_gma/////////////////////////////
        maxim = 0
        Dim yy As Double
        For i = AC_Analyze(1) To AC_Analyze(No_Sample) Step 0.005
            yy = 0
            For j = 0 To gmb_deg_ploy_txt
                yy = yy + (A_Gma(j) * (i) ^ j)
            Next
            If (yy > maxim) Then
                maxim = yy
                optimum_gma = i
            End If
        Next
        Max_Gmb = maxim
        REM /////////////////////////////////////////////////
        maxim = 0
        For i = AC_Analyze(1) To AC_Analyze(No_Sample) Step 0.005
            yy = 0
            For j = 0 To s_deg_plo_txt
                yy = yy + (A_S(j) * (i) ^ j)
            Next
            If (yy > maxim) Then
                maxim = yy
                optimum_s = i
            End If
        Next
        REM ///////////////////opimum_av/////////////////////
        maxim = 0
        For i = AC_Analyze(1) - 3 To AC_Analyze(No_Sample) + 10 Step 0.005
            yy = 0
            For j = 0 To av_deg_plo_txt
                yy = yy + (A_AV(j) * (i) ^ j)
            Next
            If (Format(yy, "0.00") = 9.0) Then
                optimum_av = i
            End If
        Next

        optimum_all = (optimum_gma + optimum_s + optimum_av) / 3.0
        op1 = optimum_gma : op2 = optimum_s : op3 = optimum_av
        optimum_Asphalt_content = optimum_all
    End Sub
    Sub draw_AIR_Voids()
        Dim F_GMB, F_S, F_F, F_AV, F_VMA As Double REM it is a factor to calculate software optimization Ratio :))
        Dim mypen As New Pen(Brushes.Green, 3)
        Dim i, yy As Double
        Dim point1, point2 As Integer
        Dim indcator As Integer REM to check result if 1 then ok if 0 then NoT oK
        indcator = 1
        REM ///////////////////Gma//////////////////////////
        Max_X_valu_draw = maxx_Gma : Min_X_valu_draw = minx_Gma : Max_Y_valu_draw = maxy_Gma : Min_Y_valu_draw = miny_Gma
        x_scale = (250.0 / (Max_X_valu_draw - Min_X_valu_draw))
        y_scale = (250.0 / (Max_Y_valu_draw - Min_Y_valu_draw))
        i = TextBox_rslt_oca.Text
        yy = 0
        For j = 0 To gmb_deg_ploy_txt
            yy = yy + (A_Gma(j) * (i) ^ j)
        Next
        x_draw = 275 + ((i - Min_X_valu_draw) * x_scale)
        y_draw = 310 - ((yy - Min_Y_valu_draw) * y_scale)
        point1 = 310 - ((0) * y_scale) + 25
        point2 = 275 + ((0) * x_scale) - 25
        Me.CreateGraphics.FillEllipse(Brushes.Green, x_draw - 4, y_draw - 4, 8, 8)
        Me.CreateGraphics.DrawLine(mypen, x_draw, point1, x_draw, y_draw)
        Me.CreateGraphics.DrawLine(mypen, point2, y_draw, x_draw, y_draw)

        TextBox_rslt_gma.Text = Format(yy, "0.000")
        F_GMB = (yy / (Max_Gmb))
        REM //////////////////////Stability//////////////////////////
        Max_X_valu_draw = maxx_S : Min_X_valu_draw = minx_S : Max_Y_valu_draw = maxy_S : Min_Y_valu_draw = miny_S
        x_scale = (250.0 / (Max_X_valu_draw - Min_X_valu_draw))
        y_scale = (250.0 / (Max_Y_valu_draw - Min_Y_valu_draw))
        i = TextBox_rslt_oca.Text
        yy = 0
        For j = 0 To s_deg_plo_txt
            yy = yy + (A_S(j) * (i) ^ j)
        Next
        x_draw = 675 + ((i - Min_X_valu_draw) * x_scale)
        y_draw = 310 - ((yy - Min_Y_valu_draw) * y_scale)
        point1 = 310 - ((0) * y_scale) + 25
        point2 = 675 + ((0) * x_scale) - 25
        Me.CreateGraphics.FillEllipse(Brushes.Green, x_draw - 4, y_draw - 4, 8, 8)
        Me.CreateGraphics.DrawLine(mypen, x_draw, point1, x_draw, y_draw)
        Me.CreateGraphics.DrawLine(mypen, point2, y_draw, x_draw, y_draw)
        TextBox_rslt_s.Text = Format(yy, "0.000")
        If (yy >= Min_stability) Then
            TextBox_rslt_s.BackColor = Color.Green
            TextBox_rslt_s.ForeColor = Color.Yellow
        Else
            TextBox_rslt_s.BackColor = Color.Red
            TextBox_rslt_s.ForeColor = Color.Yellow
            indcator = 0
        End If
        F_S = (yy - Min_stability) / (yy + Min_stability)
        REM ///////////////////////flow//////////////////////////////////
        Max_X_valu_draw = maxx_F : Min_X_valu_draw = minx_F : Max_Y_valu_draw = maxy_F : Min_Y_valu_draw = miny_F
        x_scale = (250.0 / (Max_X_valu_draw - Min_X_valu_draw))
        y_scale = (250.0 / (Max_Y_valu_draw - Min_Y_valu_draw))
        i = TextBox_rslt_oca.Text
        yy = 0
        For j = 0 To f_deg_plo_txt
            yy = yy + (A_F(j) * (i) ^ j)
        Next
        x_draw = 1075 + ((i - Min_X_valu_draw) * x_scale)
        y_draw = 310 - ((yy - Min_Y_valu_draw) * y_scale)
        point1 = 310 - ((0) * y_scale) + 25
        point2 = 1075 + ((0) * x_scale) - 25
        Me.CreateGraphics.FillEllipse(Brushes.Green, x_draw - 4, y_draw - 4, 8, 8)
        Me.CreateGraphics.DrawLine(mypen, x_draw, point1, x_draw, y_draw)
        Me.CreateGraphics.DrawLine(mypen, point2, y_draw, x_draw, y_draw)
        TextBox_rslt_f.Text = Format(yy, "0.000")
        If (yy >= flow_lower_limit And yy <= flow_upper_limit) Then
            TextBox_rslt_f.BackColor = Color.Green
            TextBox_rslt_f.ForeColor = Color.Yellow
            F_F = 1
        Else
            TextBox_rslt_f.BackColor = Color.Red
            TextBox_rslt_f.ForeColor = Color.Yellow
            indcator = 0
            F_F = 0
        End If
        REM ///////////////////////av//////////////////////////////////
        Max_X_valu_draw = maxx_AV : Min_X_valu_draw = minx_AV : Max_Y_valu_draw = maxy_AV : Min_Y_valu_draw = miny_AV
        x_scale = (250.0 / (Max_X_valu_draw - Min_X_valu_draw))
        y_scale = (250.0 / (Max_Y_valu_draw - Min_Y_valu_draw))
        i = TextBox_rslt_oca.Text
        yy = 0
        For j = 0 To av_deg_plo_txt
            yy = yy + (A_AV(j) * (i) ^ j)
        Next
        x_draw = 275 + ((i - Min_X_valu_draw) * x_scale)
        y_draw = 670 - ((yy - Min_Y_valu_draw) * y_scale)
        point1 = 670 - ((0) * y_scale) + 25
        point2 = 275 + ((0) * x_scale) - 25
        Me.CreateGraphics.FillEllipse(Brushes.Green, x_draw - 4, y_draw - 4, 8, 8)
        Me.CreateGraphics.DrawLine(mypen, x_draw, point1, x_draw, y_draw)
        Me.CreateGraphics.DrawLine(mypen, point2, y_draw, x_draw, y_draw)
        TextBox_rslt_av.Text = Format(yy, "0.000")
        If (yy >= AV_Percentage_lower_percentage And yy <= AV_Percentage_Upper_percentage) Then
            TextBox_rslt_av.BackColor = Color.Green
            TextBox_rslt_av.ForeColor = Color.Yellow
        Else
            TextBox_rslt_av.BackColor = Color.Red
            TextBox_rslt_av.ForeColor = Color.Yellow
            indcator = 0
        End If
        F_AV = 1 - Math.Abs((yy - (AV_Percentage_lower_percentage + AV_Percentage_Upper_percentage) / 2) / (yy + (AV_Percentage_lower_percentage + AV_Percentage_Upper_percentage) / 2))
        REM ///////////////////////vma//////////////////////////////////
        Max_X_valu_draw = maxx_VMA : Min_X_valu_draw = minx_VMA : Max_Y_valu_draw = maxy_VMA : Min_Y_valu_draw = miny_VMA
        x_scale = (250.0 / (Max_X_valu_draw - Min_X_valu_draw))
        y_scale = (250.0 / (Max_Y_valu_draw - Min_Y_valu_draw))
        i = TextBox_rslt_oca.Text
        yy = 0
        For j = 0 To vma_deg_plo_txt
            yy = yy + (A_VMA(j) * (i) ^ j)
        Next
        x_draw = 675 + ((i - Min_X_valu_draw) * x_scale)
        y_draw = 670 - ((yy - Min_Y_valu_draw) * y_scale)
        point1 = 670 - ((0) * y_scale) + 25
        point2 = 675 + ((0) * x_scale) - 25
        Me.CreateGraphics.FillEllipse(Brushes.Green, x_draw - 4, y_draw - 4, 8, 8)
        Me.CreateGraphics.DrawLine(mypen, x_draw, point1, x_draw, y_draw)
        Me.CreateGraphics.DrawLine(mypen, point2, y_draw, x_draw, y_draw)
        TextBox_rslt_vma.Text = Format(yy, "0.000")
        If (yy >= Min_VMA_Percentage) Then
            TextBox_rslt_vma.BackColor = Color.Green
            TextBox_rslt_vma.ForeColor = Color.Yellow
        Else
            TextBox_rslt_vma.BackColor = Color.Red
            TextBox_rslt_vma.ForeColor = Color.Yellow
            indcator = 0
        End If

        If (indcator = 1) Then
            Label_rslt_check.Text = "OK"
            Label_rslt_check.BackColor = Color.Green
            Label_rslt_check.ForeColor = Color.Yellow
        Else
            Label_rslt_check.Text = "NOT OK"
            Label_rslt_check.BackColor = Color.Red
            Label_rslt_check.ForeColor = Color.Yellow
        End If
        F_VMA = (yy - Min_VMA_Percentage) / (yy + Min_VMA_Percentage)
        REM//////////////calculation Software optimization Ratio//////////////
        If (indcator = 1) Then
            label_optimazation_ratio.Text = Format((Math.Sqrt((F_GMB ^ 2 + F_S ^ 2 + F_AV ^ 2 + F_F ^ 2 + F_VMA ^ 2) / 5)), "0.0000")
        Else
            label_optimazation_ratio.Text = Format(0, "0.0000")
        End If
    End Sub

    Private Sub Button_rslt_inc_ac_Click(sender As System.Object, e As System.EventArgs) Handles Button_rslt_inc_ac.Click
        TextBox_rslt_oca.Text += 0.05
        Dim max_AC As Double
        max_AC = 0
        For i = 1 To No_Sample
            If AC_Analyze(i) > max_AC Then max_AC = AC_Analyze(i)
        Next
        Me.CreateGraphics.Clear(BackColor)
        degree_change("gma", 0)
        If CheckBox_draw_curve.Checked = True Then draw_AIR_Voids()

        If TextBox_rslt_oca.Text >= max_AC Then TextBox_rslt_oca.Text = max_AC
    End Sub
    Private Sub Button_rslt_dec_ac_Click(sender As System.Object, e As System.EventArgs) Handles Button_rslt_dec_ac.Click
        TextBox_rslt_oca.Text -= 0.05
        Dim min_AC As Double
        min_AC = 10000000
        For i = 1 To No_Sample
            If AC_Analyze(i) < min_AC Then min_AC = AC_Analyze(i)
        Next
        Me.CreateGraphics.Clear(BackColor)
        degree_change("gma", 0)
        If CheckBox_draw_curve.Checked = True Then draw_AIR_Voids()
        If TextBox_rslt_oca.Text <= min_AC Then TextBox_rslt_oca.Text = min_AC
    End Sub

    Private Sub CheckBox_draw_curve_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox_draw_curve.CheckedChanged
        If CheckBox_draw_curve.Checked = True Then
            degree_change("gma", 0)
            draw_AIR_Voids()
        Else
            degree_change("gma", 0)
        End If

    End Sub
    Sub clear_result()
        CheckBox_draw_curve.Checked = False
        TextBox_rslt_gma.Text = ""
        TextBox_rslt_gma.BackColor = BackColor

        TextBox_rslt_s.Text = ""
        TextBox_rslt_s.BackColor = BackColor

        TextBox_rslt_f.Text = ""
        TextBox_rslt_f.BackColor = BackColor

        TextBox_rslt_av.Text = ""
        TextBox_rslt_av.BackColor = BackColor

        TextBox_rslt_vma.Text = ""
        TextBox_rslt_vma.BackColor = BackColor



    End Sub

    Private Sub Save_Screen_Main_Click(sender As System.Object, e As System.EventArgs) Handles Save_Screen_Main.Click
        My.Forms.Form1.Show()
    End Sub

   
End Class
