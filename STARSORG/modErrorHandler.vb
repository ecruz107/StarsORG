Module modErrorHandler
    'useful routine for input validation
    Public Function ValidateTextBoxLength(ByRef obj As TextBox, ByRef errP As ErrorProvider) As Boolean
        If obj.Text.Length = 0 Then
            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleLeft)
            errP.SetError(obj, "You must enter a value here")
            obj.Focus()
            Return False
        Else
            errP.SetError(obj, "")
            Return True
        End If
    End Function
    Public Function ValidateTextBoxNumeric(ByRef obj As TextBox, ByRef errP As ErrorProvider) As Boolean
        If Not IsNumeric(obj.Text) Then
            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleLeft)
            errP.SetError(obj, "You must enter a numeric value")
            obj.Focus()
            Return False
        Else
            errP.SetError(obj, "")
            Return True
        End If
    End Function

    Public Function ValidateTextBoxDate(ByRef obj As TextBox, ByRef errP As ErrorProvider) As Boolean
        If Not IsDate(obj.Text) Then

            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleLeft)
            errP.SetError(obj, "You must enter a valid date value")
            obj.Focus()
            Return False
        Else
            errP.SetError(obj, "")
            Return True
        End If

    End Function

    Public Function ValidateCombo(ByRef obj As ComboBox, ByRef errP As ErrorProvider) As Boolean
        If obj.SelectedIndex = -1 Then
            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleLeft)
            errP.SetError(obj, "You must make a selection")
            obj.Focus()
            Return False
        Else
            errP.SetError(obj, "")
            Return True



        End If
    End Function

    Public Function ValidateMaskedTextBoxLength(ByRef obj As MaskedTextBox, ByRef errP As ErrorProvider) As Boolean
        If obj.Text.Length = 0 Then
            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleLeft)
            errP.SetError(obj, "You must enter a value here")
            obj.Focus()
            Return False
        Else
            errP.SetError(obj, "")
            Return True
        End If

    End Function
End Module
