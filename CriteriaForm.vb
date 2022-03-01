Public Class CriteriaForm
    Private Sub CriteriaForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If (Degree_Pol_text.save = 0) Then
            REM *************************************************
            If (Degree_Pol_text.CheckBox_Light.Checked = True) Then
                If (Degree_Pol_text.CheckBox_N.Checked = True) Then
                    TextBox_Min_Stability.Text = "3336"
                ElseIf (Degree_Pol_text.CheckBox_Kg.Checked = True) Then
                    TextBox_Min_Stability.Text = "340.2"
                ElseIf (Degree_Pol_text.CheckBox_Ib.Checked = True) Then
                    TextBox_Min_Stability.Text = "750"
                End If
            End If
            REM *************************************************
            If (Degree_Pol_text.CheckBox_Medium.Checked = True) Then
                If (Degree_Pol_text.CheckBox_N.Checked = True) Then
                    TextBox_Min_Stability.Text = "5338"
                ElseIf (Degree_Pol_text.CheckBox_Kg.Checked = True) Then
                    TextBox_Min_Stability.Text = "544.3"
                ElseIf (Degree_Pol_text.CheckBox_Ib.Checked = True) Then
                    TextBox_Min_Stability.Text = "1200"
                End If
            End If
            REM *************************************************
            If (Degree_Pol_text.CheckBox_Heavy.Checked = True) Then
                If (Degree_Pol_text.CheckBox_N.Checked = True) Then
                    TextBox_Min_Stability.Text = "8006"
                ElseIf (Degree_Pol_text.CheckBox_Kg.Checked = True) Then
                    TextBox_Min_Stability.Text = "816.5"
                ElseIf (Degree_Pol_text.CheckBox_Ib.Checked = True) Then
                    TextBox_Min_Stability.Text = "1800"
                End If
            End If
            REM *************************************************
            If (Degree_Pol_text.CheckBox_Light.Checked = True) Then
                If (Degree_Pol_text.CheckBox_Unit.Checked = True) Then
                    TextBox_flow_lower.Text = "8" : TextBox_flow_upper.Text = "18"
                ElseIf (Degree_Pol_text.CheckBox_25mm.Checked = True) Then
                    TextBox_flow_lower.Text = "2" : TextBox_flow_upper.Text = "4.5"
                ElseIf (Degree_Pol_text.CheckBox_001in.Checked = True) Then
                    TextBox_flow_lower.Text = "0.08" : TextBox_flow_upper.Text = "0.18"
                End If
            End If
            REM *************************************************
            If (Degree_Pol_text.CheckBox_Medium.Checked = True) Then
                If (Degree_Pol_text.CheckBox_Unit.Checked = True) Then
                    TextBox_flow_lower.Text = "8" : TextBox_flow_upper.Text = "16"
                ElseIf (Degree_Pol_text.CheckBox_25mm.Checked = True) Then
                    TextBox_flow_lower.Text = "2" : TextBox_flow_upper.Text = "4"
                ElseIf (Degree_Pol_text.CheckBox_001in.Checked = True) Then
                    TextBox_flow_lower.Text = "0.08" : TextBox_flow_upper.Text = "0.16"
                End If
            End If
            REM *************************************************
            If (Degree_Pol_text.CheckBox_Heavy.Checked = True) Then
                If (Degree_Pol_text.CheckBox_Unit.Checked = True) Then
                    TextBox_flow_lower.Text = "8" : TextBox_flow_upper.Text = "14"
                ElseIf (Degree_Pol_text.CheckBox_25mm.Checked = True) Then
                    TextBox_flow_lower.Text = "2" : TextBox_flow_upper.Text = "3.5"
                ElseIf (Degree_Pol_text.CheckBox_001in.Checked = True) Then
                    TextBox_flow_lower.Text = "0.08" : TextBox_flow_upper.Text = "0.14"
                End If
            End If
            REM *************************************************
            TextBox_AV_lower.Text = "3"
            TextBox_AV_Upper.Text = "5"
            TextBox_Min_VMA.Text = "14"
            REM *************************************************
            Degree_Pol_text.Min_stability = Val(TextBox_Min_Stability.Text)
            Degree_Pol_text.flow_lower_limit = Val(TextBox_flow_lower.Text)
            Degree_Pol_text.flow_upper_limit = Val(TextBox_flow_upper.Text)
            Degree_Pol_text.AV_Percentage_lower_percentage = Val(TextBox_AV_lower.Text)
            Degree_Pol_text.AV_Percentage_Upper_percentage = Val(TextBox_AV_Upper.Text)
            Degree_Pol_text.Min_VMA_Percentage = Val(TextBox_Min_VMA.Text)
        Else
            TextBox_Min_Stability.Text = Degree_Pol_text.Min_stability
            TextBox_flow_lower.Text = Degree_Pol_text.flow_lower_limit
            TextBox_flow_upper.Text = Degree_Pol_text.flow_upper_limit
            TextBox_AV_lower.Text = Degree_Pol_text.AV_Percentage_lower_percentage
            TextBox_AV_Upper.Text = Degree_Pol_text.AV_Percentage_Upper_percentage
            TextBox_Min_VMA.Text = Degree_Pol_text.Min_VMA_Percentage
        End If
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button_save.Click
        Degree_Pol_text.Min_stability = Val(TextBox_Min_Stability.Text)
        Degree_Pol_text.flow_lower_limit = Val(TextBox_flow_lower.Text)
        Degree_Pol_text.flow_upper_limit = Val(TextBox_flow_upper.Text)
        Degree_Pol_text.AV_Percentage_lower_percentage = Val(TextBox_AV_lower.Text)
        Degree_Pol_text.AV_Percentage_Upper_percentage = Val(TextBox_AV_Upper.Text)
        Degree_Pol_text.Min_VMA_Percentage = Val(TextBox_Min_VMA.Text)
        Degree_Pol_text.save = Degree_Pol_text.save + 1
        Degree_Pol_text.Button_Run.Enabled = True
        Me.Close()
    End Sub
    Private Sub Button_Default_Click(sender As System.Object, e As System.EventArgs) Handles Button_Default.Click
        REM *************************************************
        If (Degree_Pol_text.CheckBox_Light.Checked = True) Then
            If (Degree_Pol_text.CheckBox_N.Checked = True) Then
                TextBox_Min_Stability.Text = "3336"
            ElseIf (Degree_Pol_text.CheckBox_Kg.Checked = True) Then
                TextBox_Min_Stability.Text = "340.2"
            ElseIf (Degree_Pol_text.CheckBox_Ib.Checked = True) Then
                TextBox_Min_Stability.Text = "750"
            End If
        End If
        REM *************************************************
        If (Degree_Pol_text.CheckBox_Medium.Checked = True) Then
            If (Degree_Pol_text.CheckBox_N.Checked = True) Then
                TextBox_Min_Stability.Text = "5338"
            ElseIf (Degree_Pol_text.CheckBox_Kg.Checked = True) Then
                TextBox_Min_Stability.Text = "544.3"
            ElseIf (Degree_Pol_text.CheckBox_Ib.Checked = True) Then
                TextBox_Min_Stability.Text = "1200"
            End If
        End If
        REM *************************************************
        If (Degree_Pol_text.CheckBox_Heavy.Checked = True) Then
            If (Degree_Pol_text.CheckBox_N.Checked = True) Then
                TextBox_Min_Stability.Text = "8006"
            ElseIf (Degree_Pol_text.CheckBox_Kg.Checked = True) Then
                TextBox_Min_Stability.Text = "816.5"
            ElseIf (Degree_Pol_text.CheckBox_Ib.Checked = True) Then
                TextBox_Min_Stability.Text = "1800"
            End If
        End If
        REM *************************************************
        If (Degree_Pol_text.CheckBox_Light.Checked = True) Then
            If (Degree_Pol_text.CheckBox_Unit.Checked = True) Then
                TextBox_flow_lower.Text = "8" : TextBox_flow_upper.Text = "18"
            ElseIf (Degree_Pol_text.CheckBox_25mm.Checked = True) Then
                TextBox_flow_lower.Text = "2" : TextBox_flow_upper.Text = "4.5"
            ElseIf (Degree_Pol_text.CheckBox_001in.Checked = True) Then
                TextBox_flow_lower.Text = "0.08" : TextBox_flow_upper.Text = "0.18"
            End If
        End If
        REM *************************************************
        If (Degree_Pol_text.CheckBox_Medium.Checked = True) Then
            If (Degree_Pol_text.CheckBox_Unit.Checked = True) Then
                TextBox_flow_lower.Text = "8" : TextBox_flow_upper.Text = "16"
            ElseIf (Degree_Pol_text.CheckBox_25mm.Checked = True) Then
                TextBox_flow_lower.Text = "2" : TextBox_flow_upper.Text = "4"
            ElseIf (Degree_Pol_text.CheckBox_001in.Checked = True) Then
                TextBox_flow_lower.Text = "0.08" : TextBox_flow_upper.Text = "0.16"
            End If
        End If
        REM *************************************************
        If (Degree_Pol_text.CheckBox_Heavy.Checked = True) Then
            If (Degree_Pol_text.CheckBox_Unit.Checked = True) Then
                TextBox_flow_lower.Text = "8" : TextBox_flow_upper.Text = "14"
            ElseIf (Degree_Pol_text.CheckBox_25mm.Checked = True) Then
                TextBox_flow_lower.Text = "2" : TextBox_flow_upper.Text = "3.5"
            ElseIf (Degree_Pol_text.CheckBox_001in.Checked = True) Then
                TextBox_flow_lower.Text = "0.08" : TextBox_flow_upper.Text = "0.14"
            End If
        End If
        REM *************************************************
        TextBox_AV_lower.Text = "3"
        TextBox_AV_Upper.Text = "5"
        TextBox_Min_VMA.Text = "14"
        REM *************************************************
    End Sub

    Private Sub Save_Screen_Criteria_Click(sender As System.Object, e As System.EventArgs) Handles Save_Screen_Criteria.Click
        My.Forms.Form1.Show()
    End Sub
End Class