Imports System.Data.SqlClient
Public Class CEvent_RSVP
    Private _mstrEventID As String
    Private _mstrFName As String
    Private _mstrLName As String
    Private _mstrEmail As String
    Private _isNewMember As Boolean

    Public Sub New()
        _mstrEventID = ""
        _mstrFName = ""
        _mstrLName = ""
        _mstrEmail = ""



    End Sub

#Region "Exposed Properties"

    Public Property EventID As String
        Get
            Return _mstrEventID

        End Get
        Set(strVal As String)
            _mstrEventID = strVal


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


    Public Property Email As String
        Get
            Return _mstrEmail

        End Get
        Set(strVal As String)
            _mstrEmail = strVal


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
#End Region

    Public ReadOnly Property GetSaveParameters() As ArrayList
        Get
            Dim params As New ArrayList
            params.Add(New SqlParameter("eventID", _mstrEventID))
            params.Add(New SqlParameter("fName", _mstrFName))
            params.Add(New SqlParameter("lName", _mstrLName))
            params.Add(New SqlParameter("email", _mstrEmail))

            Return params
        End Get
    End Property

    Public Function Save() As Integer
        If isNewMember Then
            Dim params As New ArrayList
            params.Add(New SqlParameter("eventID", _mstrEventID))
            Dim strResult As String = myDB.GetSingleValueFromSP("sp_CheckEventRSVPIDExist", params)
            If Not strResult = 0 Then
                Return -1 'not unique

            End If
        End If
        Return myDB.ExecSP("sp_saveEvent_RSVP", GetSaveParameters())
    End Function

    Public Function GetReportData() As SqlDataAdapter
        Return myDB.GetDataAdapterBySP("dbo.sp_getAllEvent_RSVP", Nothing)
    End Function



End Class
