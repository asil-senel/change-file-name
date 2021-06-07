Public Class Form1
    Dim sSourceLoc As String
    Dim sSourceName As String
    Dim sSourceType As String

    Dim sTargetName As String
    Dim sTargetType As String

    Dim iDiffUp As Integer
    Dim iDiffDown As Integer
    Dim iDiff As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        sSourceLoc = TextBox1.Text
        sSourceName = TextBox3.Text
        sSourceType = TextBox5.Text

        If TextBox4.Text = "" Then
            sTargetName = TextBox3.Text
        Else
            sTargetName = TextBox4.Text
        End If

        If TextBox6.Text = "" Then
            sTargetType = TextBox5.Text
        Else
            sTargetType = TextBox6.Text
        End If

        If ((TextBox7.Text <> "") Or (TextBox8.Text <> "") Or (TextBox9.Text <> "") Or (TextBox10.Text <> "")) Then
            iDiffUp = CInt(TextBox10.Text) - CInt(TextBox8.Text)
            iDiffDown = CInt(TextBox9.Text) - CInt(TextBox7.Text)
        End If

        If iDiffUp = iDiffDown Then
            iDiff = iDiffDown
        End If

        If ((TextBox7.Text <> "") Or (TextBox8.Text <> "")) Then 'index on for both
            If ((CheckBox1.Checked) = True And (CheckBox2.Checked = False)) Then 'source parentheses exist, target parentheses don't exist
                For rangeS As Integer = TextBox7.Text To TextBox8.Text
                    My.Computer.FileSystem.RenameFile(sSourceLoc + sSourceName + "(" + CStr(rangeS) + ")" + "." + sSourceType, sTargetName + CStr(rangeS + iDiff) + "." + sTargetType)
                Next
            ElseIf ((CheckBox1.Checked) = True And (CheckBox2.Checked = True)) Then 'both source and target parentheses exist
                For rangeS As Integer = TextBox7.Text To TextBox8.Text
                    My.Computer.FileSystem.RenameFile(sSourceLoc + sSourceName + "(" + CStr(rangeS) + ")" + "." + sSourceType, sTargetName + "(" + CStr(rangeS + iDiff) + ")" + "." + sTargetType)
                Next
            ElseIf ((CheckBox1.Checked) = False And (CheckBox2.Checked = True)) Then 'source parentheses don't exist, target parentheses do
                For rangeS As Integer = TextBox7.Text To TextBox8.Text
                    My.Computer.FileSystem.RenameFile(sSourceLoc + sSourceName + CStr(rangeS) + "." + sSourceType, sTargetName + "(" + CStr(rangeS + iDiff) + ")" + "." + sTargetType)
                Next
            Else 'both source and target parentheses don't exist
                For rangeS As Integer = TextBox7.Text To TextBox8.Text
                    My.Computer.FileSystem.RenameFile(sSourceLoc + sSourceName + CStr(rangeS) + "." + sSourceType, sTargetName + CStr(rangeS + iDiff) + "." + sTargetType)
                Next
            End If
        ElseIf ((TextBox7.Text <> "") And (TextBox8.Text <> "") And (TextBox9.Text = "") And (TextBox10.Text = "")) Then 'index on for source
            If ((CheckBox1.Checked) = True And (CheckBox2.Checked = False)) Then 'source parentheses exist, target parentheses don't exist
                For rangeS As Integer = TextBox7.Text To TextBox8.Text
                    My.Computer.FileSystem.RenameFile(sSourceLoc + sSourceName + "(" + CStr(rangeS) + ")" + "." + sSourceType, sTargetName + "." + sTargetType)
                Next
            ElseIf ((CheckBox1.Checked) = True And (CheckBox2.Checked = True)) Then 'both source and target parentheses exist
                For rangeS As Integer = TextBox7.Text To TextBox8.Text
                    My.Computer.FileSystem.RenameFile(sSourceLoc + sSourceName + "(" + CStr(rangeS) + ")" + "." + sSourceType, sTargetName + "." + sTargetType)
                Next
            ElseIf ((CheckBox1.Checked) = False And (CheckBox2.Checked = True)) Then 'source parentheses don't exist, target parentheses do
                For rangeS As Integer = TextBox7.Text To TextBox8.Text
                    My.Computer.FileSystem.RenameFile(sSourceLoc + sSourceName + CStr(rangeS) + "." + sSourceType, sTargetName + "." + sTargetType)
                Next
            Else 'both source and target parentheses don't exist
                For rangeS As Integer = TextBox7.Text To TextBox8.Text
                    My.Computer.FileSystem.RenameFile(sSourceLoc + sSourceName + CStr(rangeS) + "." + sSourceType, sTargetName + CStr(rangeS + iDiff) + "." + sTargetType)
                Next
            End If
        ElseIf ((TextBox7.Text = "") And (TextBox8.Text = "") And (TextBox9.Text <> "") And (TextBox10.Text <> "")) Then 'index on for target
            If ((CheckBox1.Checked) = True And (CheckBox2.Checked = False)) Then 'source parentheses exist, target parentheses don't exist
                For rangeS As Integer = TextBox9.Text To TextBox10.Text
                    My.Computer.FileSystem.RenameFile(sSourceLoc + sSourceName + "." + sSourceType, sTargetName + CStr(rangeS + iDiff) + "." + sTargetType)
                Next
            ElseIf ((CheckBox1.Checked) = True And (CheckBox2.Checked = True)) Then 'both source and target parentheses exist
                For rangeS As Integer = TextBox9.Text To TextBox10.Text
                    My.Computer.FileSystem.RenameFile(sSourceLoc + sSourceName + "." + sSourceType, sTargetName + "(" + CStr(rangeS + iDiff) + ")" + "." + sTargetType)
                Next
            ElseIf ((CheckBox1.Checked) = False And (CheckBox2.Checked = True)) Then 'source parentheses don't exist, target parentheses do
                For rangeS As Integer = TextBox9.Text To TextBox10.Text
                    My.Computer.FileSystem.RenameFile(sSourceLoc + sSourceName + "." + sSourceType, sTargetName + "(" + CStr(rangeS + iDiff) + ")" + "." + sTargetType)
                Next
            Else 'both source and target parentheses don't exist
                For rangeS As Integer = TextBox9.Text To TextBox10.Text
                    My.Computer.FileSystem.RenameFile(sSourceLoc + sSourceName + "." + sSourceType, sTargetName + CStr(rangeS + iDiff) + "." + sTargetType)
                Next
            End If
        ElseIf ((TextBox7.Text = "") And (TextBox8.Text = "") And (TextBox9.Text = "") And (TextBox10.Text = "")) Then 'index off for both
            My.Computer.FileSystem.RenameFile(sSourceLoc + sSourceName + "." + sSourceType, sTargetName + "." + sTargetType)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        CheckBox1.CheckState = CheckState.Unchecked
        CheckBox2.CheckState = CheckState.Unchecked
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = ""
    End Sub
End Class
