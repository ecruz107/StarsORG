﻿Imports System.Data.SqlClient
Imports Microsoft.Reporting.WinForms


Public Class frmRolesReport
    Private ds As DataSet
    Private da As SqlDataAdapter
    Private Role As CRole
    Private Sub frmRolesReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.rpvRoleReport.RefreshReport()
    End Sub

    Public Sub Display()
        Role = New CRole
        rpvRoleReport.LocalReport.ReportPath = AppDomain.CurrentDomain.BaseDirectory & "Reports\rptRoles.rdlc"
        ds = New DataSet
        da = Role.GetReportData()
        da.Fill(ds)
        rpvRoleReport.LocalReport.DataSources.Add(New ReportDataSource("dsRoles1", ds.Tables(0)))
        rpvRoleReport.SetDisplayMode(DisplayMode.PrintLayout)
        rpvRoleReport.RefreshReport()
        Me.Cursor = Cursors.Default
        Me.ShowDialog()

    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

        Me.Close()

    End Sub
End Class