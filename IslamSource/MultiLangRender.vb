Public Class MultiLangRender
    Dim _RenderArray As IslamMetadata.RenderArray
    Dim Texts As New Generic.List(Of Windows.Forms.TextBox)
    Dim MultiLangRenders As Generic.List(Of MultiLangRender)
    Public Property RenderArray
        Get
            Return _RenderArray
        End Get
        Set(value)
            _RenderArray = value
        End Set
    End Property

    Private Sub RecalcLayout()
        If _RenderArray Is Nothing Then Return
        Dim Top As Integer = 0
        Dim Right As Integer = MaximumSize.Width
        Dim MaxRight As Integer = MaximumSize.Width
        Me.RightToLeft = Windows.Forms.RightToLeft.Yes
        For Count As Integer = 0 To _RenderArray.Items.Count - 1
            For SubCount As Integer = 0 To _RenderArray.Items(Count).TextItems.Length - 1
                If _RenderArray.Items(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eArabic Then
                    Dim NewText As New Windows.Forms.TextBox
                    NewText.RightToLeft = Windows.Forms.RightToLeft.Yes
                    NewText.Text = _RenderArray.Items(Count).TextItems(SubCount).Text
                    'NewText.Font.Size = 2
                    Texts.Add(NewText)
                    Me.Controls.Add(NewText)
                    NewText.Width = NewText.PreferredSize.Width
                    If Right = NewText.PreferredSize.Width < 0 Then
                        NewText.Left = 0
                        Top += Texts(0).PreferredSize.Height
                        MaxRight = Math.Min(Right, MaxRight)
                        Right = NewText.PreferredSize.Width
                    Else
                        NewText.Left = Right
                        Right += NewText.PreferredSize.Width
                    End If
                    NewText.Top = Top
                End If
            Next
        Next
        Top += Texts(0).PreferredSize.Height
        MaxRight = Math.Min(Right, MaxRight)
        Me.Size = New Size(MaxRight, Top)
    End Sub

    Private Sub MultiLangRender_Load(sender As Object, e As EventArgs) Handles Me.Load
        RecalcLayout()
    End Sub
    Private Sub MultiLangRender_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        RecalcLayout()
    End Sub
End Class
