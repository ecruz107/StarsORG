Imports System.Data.SqlClient
Public Class CMembers
    Private _Member As CMember

    Public Sub New()
        _Member = New CMember
    End Sub

    Public ReadOnly Property currentObject() As CMember
        Get
            Return _Member
        End Get
    End Property
    Public Sub Clear()
        _Member = New CMember
    End Sub

    Public Sub CreateNewMember()
        Clear()
        _Member.isNewMember = True
    End Sub

    Public Function Save() As Integer
        Return _Member.Save()

    End Function

    Public Function GetAllMembers() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllMembers", Nothing)
        Return objDR
    End Function

    Public Function GetMemberByPID(strPID As String) As CMember
        Dim params As New ArrayList
        Dim strArr() As String = strPID.Split(" ")

        params.Add(New SqlParameter("pid", strArr(1)))

        FillObject(myDB.GetDataReaderBySP("sp_getMemberByPID", params))

        Return _Member
    End Function

    Private Function FillObject(objDR As SqlDataReader) As CMember
        If objDR.Read Then
            With _Member
                .PID = objDR.Item("PID")
                .FName = objDR.Item("FName")
                .LName = objDR.Item("LName")
                .MI = objDR.Item("MI")
                .Email = objDR.Item("Email")
                .Phone = objDR.Item("Phone")
                .PhotoPath = objDR.Item("PhotoPath")
            End With
        End If
        objDR.Close()
        Return _Member
    End Function

End Class
