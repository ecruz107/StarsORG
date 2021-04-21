Imports System.Data.SqlClient

Public Class CRoles
    Private _Role As CRole

    Public Sub New()
        _Role = New CRole
    End Sub

    Public ReadOnly Property currentObject() As CRole
        Get
            Return _Role


        End Get
    End Property

    Public Sub Clear()
        _Role = New CRole
    End Sub

    Public Sub CreateNewRole()
        Clear()
        _Role.IsNewRole = True
    End Sub
    Public Function Save() As Integer
        Return _Role.Save()
    End Function

    Public Function GetAllRoles() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllRoles", Nothing)
        Return objDR
    End Function
    Public Function GetRoleByRoleID(strID As String) As CRole
        Dim params As New ArrayList
        Dim objDR As SqlDataReader
        params.Add(New SqlParameter("roleID", strID))
        'objDR = myDB.GetDataReaderBySP("sp_getRoleByRoleID", params)
        'Return objDR
        FillObject(myDB.GetDataReaderBySP("sp_getRoleByRoleID", params))
        Return _Role
    End Function

    Private Function FillObject(objDR As SqlDataReader) As CRole
        If objDR.Read Then
            With _Role
                .RoleID = objDR.Item("RoleID")
                .RoleDescription = objDR.Item("RoleDescription")
            End With
        Else 'do nothing
        End If
        objDR.Close()
        Return _Role
    End Function


End Class
