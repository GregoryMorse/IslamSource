Public Class MultiLangRender
    Dim MainText As New System.Windows.Forms.TextBox
    Public Property ArabicText
        Get
            Return MainText.Text
        End Get
        Set(value)
            MainText.Text = value
        End Set
    End Property


    Private Sub MultiLangRender_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Controls.Add(MainText)
    End Sub
End Class
