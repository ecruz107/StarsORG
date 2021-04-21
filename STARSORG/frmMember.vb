Imports System.Data.SqlClient
Public Class frmMember
    Private objMembers As CMembers
    Private blnClearing As Boolean
    Private blnReloading As Boolean





#Region "ToolBar Stuff"
    Private Sub tsbProxy_MouseEnter(sender As Object, e As EventArgs) Handles tsbCourse.MouseEnter, tsbEvent.MouseEnter, tsbHelp.MouseEnter, tsbHome.MouseEnter, tsbLogOut.MouseEnter, tsbMember.MouseEnter, tsbRole.MouseEnter, tsbRSVP.MouseEnter, tsbSemester.MouseEnter, tsbTutor.MouseEnter
        Dim tsbProxy As ToolStripButton
        tsbProxy = DirectCast(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Text
    End Sub

    Private Sub tsbProxy_MouseLeave(sender As Object, e As EventArgs) Handles tsbCourse.MouseLeave, tsbEvent.MouseLeave, tsbHelp.MouseLeave, tsbHome.MouseLeave, tsbLogOut.MouseLeave, tsbMember.MouseLeave, tsbRole.MouseLeave, tsbRSVP.MouseLeave, tsbSemester.MouseLeave, tsbTutor.MouseLeave
        Dim tsbProxy As ToolStripButton
        tsbProxy = DirectCast(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Image
    End Sub


    Private Sub tsbCourse_Click(sender As Object, e As EventArgs) Handles tsbCourse.Click
        intNextAction = ACTION_COURSE
        Me.Hide()

    End Sub

    Private Sub tsbEvent_Click(sender As Object, e As EventArgs) Handles tsbEvent.Click
        intNextAction = ACTION_EVENT
        Me.Hide()

    End Sub

    Private Sub tsbHelp_Click(sender As Object, e As EventArgs) Handles tsbHelp.Click
        intNextAction = ACTION_HELP
        Me.Hide()

    End Sub

    Private Sub tsbHome_Click(sender As Object, e As EventArgs) Handles tsbHome.Click
        intNextAction = ACTION_HOME
        Me.Hide()

    End Sub

    Private Sub tsbLogOut_Click(sender As Object, e As EventArgs) Handles tsbLogOut.Click
        intNextAction = ACTION_LOGOUT
        Me.Hide()

    End Sub

    Private Sub tsbMember_Click(sender As Object, e As EventArgs) Handles tsbMember.Click
        'Do nothing

    End Sub

    Private Sub tsbRole_Click(sender As Object, e As EventArgs) Handles tsbRole.Click
        intNextAction = ACTION_ROLE
        Me.Hide()


    End Sub

    Private Sub tsbRSVP_Click(sender As Object, e As EventArgs) Handles tsbRSVP.Click
        intNextAction = ACTION_RSVP
        Me.Hide()

    End Sub

    Private Sub tsbSemester_Click(sender As Object, e As EventArgs) Handles tsbSemester.Click
        intNextAction = ACTION_SEMESTER
        Me.Hide()

    End Sub

    Private Sub tsbTutor_Click(sender As Object, e As EventArgs) Handles tsbTutor.Click
        intNextAction = ACTION_TUTOR
        Me.Hide()

    End Sub
#End Region


    Private Sub LoadMembers()
        Dim objDR As SqlDataReader

        lstMember.Items.Clear()
        Try
            objDR = objMembers.GetAllMembers

            Do While objDR.Read

                lstMember.Items.Add(objDR.Item("LName") + " " + objDR.Item("PID"))


                lstMember.Sorted = True



            Loop



            objDR.Close()


        Catch ex As Exception

        End Try
        If objMembers.currentObject.PID <> "" Then
            lstMember.SelectedIndex = lstMember.FindStringExact(objMembers.currentObject.PID)
        End If
        blnReloading = False
    End Sub

    Private Sub frmMember_Load(sender As Object, e As EventArgs) Handles Me.Load
        objMembers = New CMembers

    End Sub

    Private Sub frmMember_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ClearScreenControls(Me)
        LoadMembers()
        grpEdit.Enabled = False
        grpPhoto.Hide()

    End Sub

    Private Sub lstMembers_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstMember.SelectedIndexChanged
        If blnClearing Then
            Exit Sub
        End If
        If blnReloading Then
            Exit Sub
        End If
        If lstMember.SelectedIndex = -1 Then
            Exit Sub
        End If


        chkNew.Checked = False
        LoadSelectedRecord()

        grpEdit.Enabled = True


    End Sub

    Private Sub LoadSelectedRecord()

        Try
            objMembers.GetMemberByPID(lstMember.SelectedItem.ToString)




            With objMembers.currentObject

                txtPID.Text = .PID
                txtFName.Text = .FName
                txtLName.Text = .LName
                txtEmail.Text = .Email
                txtMI.Text = .MI
                txtPhone.Text = .Phone
                If .PhotoPath = "" Then
                    pbPhotoPath.BackgroundImage = Nothing
                    pbPhotoPath.Hide()
                    grpPhoto.Show()


                Else
                    pbPhotoPath.BackgroundImage = Image.FromFile(.PhotoPath)
                    pbPhotoPath.Show()
                    grpPhoto.Hide()
                End If

            End With

        Catch ex As Exception

            MessageBox.Show("Error loading members values: " & ex.ToString, "Program error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub chkNew_CheckedChanged(sender As Object, e As EventArgs) Handles chkNew.CheckedChanged
        If blnClearing Then
            Exit Sub
        End If

        If chkNew.Checked Then
            tslStatus.text = ""
            txtPID.Clear()
            txtFName.Clear()
            txtLName.Clear()
            txtMI.Clear()
            txtPhone.Clear()
            txtEmail.Clear()



            pbPhotoPath.BackgroundImage = Nothing
            pbPhotoPath.Hide()

            grpPhoto.Show()





            lstMember.SelectedIndex = -1
            grpMembers.Enabled = False
            grpEdit.Enabled = True
            txtPID.Focus()
            objMembers.currentObject.isNewMember = False
        Else
            grpMembers.Enabled = True
            grpEdit.Enabled = False
            objMembers.currentObject.isNewMember = False
            pbPhotoPath.Show()
            grpPhoto.Hide()
        End If
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged


        lstMember.SelectedIndex = lstMember.FindString(txtSearch.Text)


        If txtSearch.Text = "" Then
            lstMember.SelectedIndex = -1
        End If

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim intResult As Integer
        Dim blnErrors As Boolean
        tslStatus.Text = ""
        errP.Clear()

        If Not ValidateTextBoxLength(txtFName, errP) Then
            blnErrors = True
        End If
        If Not ValidateTextBoxLength(txtLName, errP) Then
            blnErrors = True
        End If
        If Not IsValidEmailAddress(txtEmail.Text) Then
            blnErrors = True
            errP.SetError(txtEmail, "Please enter a valid email")
        End If
        If Not ValidateTextBoxLength(txtPID, errP) Then
            blnErrors = True
        End If
        If Not ValidatePhotoPath() Then
            blnErrors = True
            errP.SetError(txtPhoto, "Photo path does not exist")
        End If

        If blnErrors Then
            Exit Sub
        End If

        With objMembers.currentObject
            .PID = txtPID.Text
            .FName = txtFName.Text
            .LName = txtLName.Text
            .MI = txtMI.Text
            .Phone = txtPhone.Text
            .Email = txtEmail.Text
            .PhotoPath = txtPhoto.Text
        End With

        Try
            Me.Cursor = Cursors.WaitCursor
            intResult = objMembers.Save
            If intResult = 1 Then
                tslStatus.Text = "Member record saved"
            End If
            If intResult = -1 Then
                MessageBox.Show("PID must be unique. Unable to save Member record", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tslStatus.Text = "Error"
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to save Member record" & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tslStatus.Text = "Error"

        End Try
        Me.Cursor = Cursors.Default
        blnReloading = True
        LoadMembers()
        chkNew.Checked = False
        grpMembers.Enabled = True
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        blnClearing = True
        tslStatus.Text = ""
        chkNew.Checked = False
        errP.Clear()
        If Not lstMember.SelectedIndex = -1 Then
            LoadSelectedRecord()
        Else
            grpEdit.Enabled = False

        End If
        blnClearing = False
        objMembers.currentObject.isNewMember = False
        grpMembers.Enabled = True
    End Sub

    Private Sub btnReport_Click(sender As Object, e As EventArgs) Handles btnReport.Click
        Dim RoleReport As New frmMemberReport
        If lstMember.Items.Count = 0 Then
            MessageBox.Show("There are no records to print")
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        RoleReport.Display()
        Me.Cursor = Cursors.Default


    End Sub
    Private Function ValidatePhotoPath() As Boolean
        Dim blnPhotoExists = False

        If My.Computer.FileSystem.FileExists(txtPhoto.Text) Or txtPhoto.Text = "" Then
            blnPhotoExists = True
        End If
        Return blnPhotoExists
    End Function

    Function IsValidEmailAddress(emailAddress As String) As Boolean
        Dim valid As Boolean = True

        Try
            If emailAddress = "" Then
                valid = False
            Else
                Dim a = New Net.Mail.MailAddress(emailAddress)
            End If

        Catch ex As FormatException
            valid = False
        End Try
        Return valid
    End Function

End Class