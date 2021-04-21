Imports System.Data.SqlClient
Public Class CMember
    Private _mstrPID As String
    Private _mstrFName As String
    Private _mstrLName As String
    Private _mstrMI As String
    Private _mstrEmail As String
    Private _mstrPhone As String
    Private _mstrPhotoPath As String
    Private _isNewMember As Boolean

    Public Sub New()
        _mstrPID = ""
        _mstrFName = ""
        _mstrLName = ""
        _mstrMI = ""
        _mstrEmail = ""
        _mstrPhone = ""
        _mstrPhotoPath = ""

    End Sub

#Region "Exposed Properties"

    Public Property PID As String
        Get
            Return _mstrPID

        End Get
        Set(strVal As String)
            _mstrPID = strVal


        End Set
    End Property

    Public Property FName As String
        Get
            Return _mstrFName

        End Get
        Set(strVal As String)
            _mstrFName = strVal


        End Set
    End Property

    Public Property LName As String
        Get
            Return _mstrLName

        End Get
        Set(strVal As String)
            _mstrLName = strVal


        End Set
    End Property

    Public Property MI As String
        Get
            Return _mstrMI

        End Get
        Set(strVal As String)
            _mstrMI = strVal


        End Set
    End Property

    Public Property Email As String
        Get
            Return _mstrEmail

        End Get
        Set(strVal As String)
            _mstrEmail = strVal


        End Set
    End Property

    Public Property Phone As String
        Get
            Return _mstrPhone

        End Get
        Set(strVal As String)
            _mstrPhone = strVal


        End Set
    End Property

    Public Property PhotoPath As String
        Get
            Return _mstrPhotoPath

        End Get
        Set(strVal As String)
            _mstrPhotoPath = strVal


        End Set
    End Property
    Public Property isNewMember As Boolean
        Get
            Return _isNewMember

        End Get
        Set(blnVal As Boolean)
            _isNewMember = blnVal


        End Set
    End Property


    Public ReadOnly Property GetSaveParameters() As ArrayList
        Get
            Dim params As New ArrayList
            params.Add(New SqlParameter("pid", _mstrPID))
            params.Add(New SqlParameter("fName", _mstrFName))
            params.Add(New SqlParameter("lName", _mstrLName))
            params.Add(New SqlParameter("mi", _mstrMI))
            params.Add(New SqlParameter("email", _mstrEmail))
            params.Add(New SqlParameter("phone", _mstrPhone))
            params.Add(New SqlParameter("photoPath", _mstrPhotoPath))
            Return params
        End Get
    End Property


#End Region

    Public Function Save() As Integer
        If isNewMember Then
            Dim params As New ArrayList
            params.Add(New SqlParameter("pid", _mstrPID))
            Dim strResult As String = myDB.GetSingleValueFromSP("sp_CheckMemberPIDExist", params)
            If Not strResult = 0 Then
                Return -1 'not unique

            End If
        End If
        Return myDB.ExecSP("sp_saveMember", GetSaveParameters())
    End Function

    Public Function GetReportData() As SqlDataAdapter
        Return myDB.GetDataAdapterBySP("dbo.sp_getAllMembers", Nothing)
    End Function
End Class
