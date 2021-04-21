Module modObjects
    Public Sub ClearScreenControls(ByVal objContainer As Control)

        Dim obj As Control
        Dim strControlType As String
        For Each obj In objContainer.Controls
            strControlType = TypeName(obj)
            Select Case strControlType
                Case "TextBox"
                    Dim cntrl As TextBox
                    cntrl = DirectCast(obj, TextBox)
                    cntrl.Clear()
                Case "CheckBox"
                    Dim cntrl As CheckBox
                    cntrl = DirectCast(obj, CheckBox)
                    cntrl.Checked = False
                Case "ComboBox"
                    Dim cntrl As ComboBox
                    cntrl = DirectCast(obj, ComboBox)
                    cntrl.SelectedIndex = -1
                Case "RadioButton"
                    Dim cntrl As RadioButton
                    cntrl = DirectCast(obj, RadioButton)
                    cntrl.Checked = False
                Case "MaskedTextBox"
                    Dim cntrl As MaskedTextBox
                    cntrl = DirectCast(obj, MaskedTextBox)
                    cntrl.Clear()
                Case "GroupBox"
                    Dim cntrl As GroupBox
                    cntrl = DirectCast(obj, GroupBox)
                    ClearScreenControls(cntrl)
                Case Else


            End Select
        Next
    End Sub
End Module
