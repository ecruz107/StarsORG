<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMemberReport
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim ReportDataSource1 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Me.rpvMemberReport = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'rpvMemberReport
        '
        ReportDataSource1.Name = "dsMember"
        ReportDataSource1.Value = Nothing
        Me.rpvMemberReport.LocalReport.DataSources.Add(ReportDataSource1)
        Me.rpvMemberReport.LocalReport.ReportEmbeddedResource = "STARSORG.rptMembers.rdlc"
        Me.rpvMemberReport.Location = New System.Drawing.Point(12, 9)
        Me.rpvMemberReport.Name = "rpvMemberReport"
        Me.rpvMemberReport.ServerReport.BearerToken = Nothing
        Me.rpvMemberReport.Size = New System.Drawing.Size(703, 368)
        Me.rpvMemberReport.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(614, 401)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(101, 37)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'frmMemberReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(727, 450)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.rpvMemberReport)
        Me.Name = "frmMemberReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmMemberReport"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents rpvMemberReport As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents btnClose As Button
End Class
