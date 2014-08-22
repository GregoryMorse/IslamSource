Public Class MultiLangRender
    <Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="GetSystemMetrics")> _
    Private Shared Function GetSystemMetrics(ByVal nIndex As Integer) As Integer
    End Function
    <Runtime.InteropServices.DllImport("gdi32.dll", EntryPoint:="SelectObject")> _
    Private Shared Function SelectObject(ByVal hdc As IntPtr, ByVal hObject As IntPtr) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("gdi32.dll", EntryPoint:="GetTextExtentExPointW")> _
    Private Shared Function GetTextExtentExPoint(ByVal hdc As IntPtr, <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.LPWStr)> ByVal lpszStr As String, ByVal cchString As Integer, ByVal nMaxExtent As Integer, ByRef lpnFit As Integer, ByVal alpDx As Integer(), ByRef lpSize As Size) As Boolean
    End Function
    Const SM_CXBORDER As Integer = 5
    Const SM_CYBORDER As Integer = 6
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
                    Dim theText As String = _RenderArray.Items(Count).TextItems(SubCount).Text
                    While theText <> String.Empty
                        Dim NewText As New Windows.Forms.TextBox
                        NewText.RightToLeft = Windows.Forms.RightToLeft.Yes
                        NewText.Font = New System.Drawing.Font(NewText.Font.FontFamily.Name, 22, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        Dim s As Drawing.Size
                        Dim nChar As Integer
                        Dim g As Graphics = NewText.CreateGraphics()
                        Dim hdc As IntPtr = g.GetHdc()
                        Dim oldFont As IntPtr = SelectObject(hdc, NewText.Font.ToHfont())
                        GetTextExtentExPoint(hdc, theText, theText.Length, MaximumSize.Width - GetSystemMetrics(SM_CXBORDER) * 2, nChar, Nothing, s)
                        SelectObject(hdc, oldFont)
                        g.ReleaseHdc(hdc)
                        g.Dispose()
                        NewText.Text = theText.Substring(0, nChar)
                        theText = theText.Substring(nChar)
                        Texts.Add(NewText)
                        Me.Controls.Add(NewText)
                        NewText.Width = NewText.PreferredSize.Width
                        If Right - NewText.PreferredSize.Width < 0 Then
                            Top += Texts(0).PreferredSize.Height
                            MaxRight = Math.Min(Right, MaxRight)
                            Right = MaximumSize.Width - NewText.PreferredSize.Width
                        Else
                            Right -= NewText.PreferredSize.Width
                        End If
                        NewText.Left = Right
                        NewText.Top = Top
                    End While
                End If
            Next
        Next
        Top += Texts(0).PreferredSize.Height
        MaxRight = MaximumSize.Width ' - Math.Min(Right, MaxRight)
        Me.Size = New Size(MaxRight, Top)
    End Sub

    Private Sub MultiLangRender_Load(sender As Object, e As EventArgs) Handles Me.Load
        RecalcLayout()
    End Sub
    Private Sub MultiLangRender_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        RecalcLayout()
    End Sub
End Class
