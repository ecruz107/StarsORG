Imports System.Data.SqlClient
Public Class CEvent_RSVPS

    Private _event_RSVP As CEvent_RSVP
    Public Sub New()
        _event_RSVP = New CEvent_RSVP
    End Sub

    Public ReadOnly Property currentObject() As CEvent_RSVP
        Get
            Return _event_RSVP
        End Get
    End Property
    Public Sub Clear()
        _event_RSVP = New CEvent_RSVP
    End Sub

    Public Sub CreateNewMember()
        Clear()
        _event_RSVP.isNewMember = True
    End Sub

    Public Function Save() As Integer
        Return _event_RSVP.Save()

    End Function

    Public Function GetAllEvent_RSVP() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllEvent_RSVP", Nothing)
        Return objDR
    End Function

    Public Function GetRSVPByEventRSVPID(strRSVPID As String) As CEvent_RSVP
        Dim params As New ArrayList


        params.Add(New SqlParameter("eventID", strRSVPID))

        FillObject(myDB.GetDataReaderBySP("sp_getRSVPByEventRSVPID", params))

        Return _event_RSVP
    End Function

    Private Function FillObject(objDR As SqlDataReader) As CEvent_RSVP
        If objDR.Read Then
            With _event_RSVP
                .EventID = objDR.Item("EventID")
                .FName = objDR.Item("FName")
                .LName = objDR.Item("LName")
                .Email = objDR.Item("Email")

            End With
        End If
        objDR.Close()
        Return _event_RSVP
    End Function
End Class
