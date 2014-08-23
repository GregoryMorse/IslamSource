Public Class MultiLangRender
    <Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="GetSystemMetrics")> _
    Private Shared Function GetSystemMetrics(ByVal nIndex As Integer) As Integer
    End Function
    <Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SendMessage")> _
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("gdi32.dll", EntryPoint:="SelectObject")> _
    Private Shared Function SelectObject(ByVal hdc As IntPtr, ByVal hObject As IntPtr) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("gdi32.dll", EntryPoint:="GetTextExtentExPointW")> _
    Private Shared Function GetTextExtentExPoint(ByVal hdc As IntPtr, <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.LPWStr)> ByVal lpszStr As String, ByVal cchString As Integer, ByVal nMaxExtent As Integer, ByRef lpnFit As Integer, ByVal alpDx As Integer(), ByRef lpSize As Size) As Boolean
    End Function
    Const SM_CXBORDER As Integer = 5
    Const SM_CYBORDER As Integer = 6
    Const EM_GETMARGINS As UInteger = &HD4
    Dim _RenderArray As IslamMetadata.RenderArray
    Dim Texts As New Generic.List(Of Generic.List(Of Generic.List(Of TextBox)))
    Public Property RenderArray
        Get
            Return _RenderArray
        End Get
        Set(value)
            _RenderArray = value
        End Set
    End Property
    Private Sub RecalcLayout()
        Me.Controls.Clear()
        Texts.Clear()
        If _RenderArray Is Nothing Then Return
        Dim Top As Integer = 0
        Dim CurTop As Integer
        Dim Right As Integer = MaximumSize.Width
        Dim MaxRight As Integer = MaximumSize.Width
        Dim NextRight As Integer
        Dim IsOverflow As Boolean
        Me.RightToLeft = Windows.Forms.RightToLeft.Yes
        For Count As Integer = 0 To _RenderArray.Items.Count - 1
            Texts.Add(New Generic.List(Of Generic.List(Of TextBox)))
            NextRight = Right
            IsOverflow = False
            Dim IsFirst As Boolean = False
            For SubCount As Integer = 0 To _RenderArray.Items(Count).TextItems.Length - 1
                Texts(Count).Add(New Generic.List(Of TextBox))
                If _RenderArray.Items(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eArabic Or _RenderArray.Items(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eLTR Or _RenderArray.Items(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eRTL Then
                    Dim theText As String = _RenderArray.Items(Count).TextItems(SubCount).Text
                    Right = NextRight
                    While theText <> String.Empty
                        Dim NewText As New TextBox
                        NewText.RightToLeft = If(_RenderArray.Items(Count).TextItems(SubCount).DisplayClass <> IslamMetadata.RenderArray.RenderDisplayClass.eLTR, Windows.Forms.RightToLeft.Yes, Windows.Forms.RightToLeft.No)
                        NewText.Font = New System.Drawing.Font(NewText.Font.FontFamily.Name, 22, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        Dim s As Drawing.Size
                        Dim nChar As Integer
                        Dim g As Graphics = NewText.CreateGraphics()
                        Dim hdc As IntPtr = g.GetHdc()
                        Dim oldFont As IntPtr = SelectObject(hdc, NewText.Font.ToHfont())
                        Dim ret As IntPtr = SendMessage(NewText.Handle, EM_GETMARGINS, IntPtr.Zero, IntPtr.Zero)
                        GetTextExtentExPoint(hdc, theText, theText.Length, MaximumSize.Width - (ret.ToInt32() And &HFFFF) - (ret.ToInt32() << 16) - GetSystemMetrics(SM_CXBORDER) * 2 - NewText.Margin.Left - NewText.Margin.Right, nChar, Nothing, s)
                        If Not IsFirst Then
                            If CurTop <> 0 Then
                                If theText.Length <> nChar Or Right - s.Width < 0 Then
                                    Top += CurTop
                                    NextRight = MaximumSize.Width
                                    Right = NextRight
                                Else
                                    For TestCount As Integer = 1 To _RenderArray.Items(Count).TextItems.Length - 1
                                        Dim tnChar As Integer
                                        Dim ts As Drawing.Size
                                        If _RenderArray.Items(Count).TextItems(TestCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eArabic Or _RenderArray.Items(Count).TextItems(TestCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eLTR Or _RenderArray.Items(Count).TextItems(TestCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eRTL Then
                                            GetTextExtentExPoint(hdc, _RenderArray.Items(Count).TextItems(TestCount).Text, _RenderArray.Items(Count).TextItems(TestCount).Text.Length, MaximumSize.Width - (ret.ToInt32() And &HFFFF) - (ret.ToInt32() << 16) - GetSystemMetrics(SM_CXBORDER) * 2 - NewText.Margin.Left - NewText.Margin.Right, tnChar, Nothing, ts)
                                            If _RenderArray.Items(Count).TextItems(TestCount).Text.Length <> tnChar Or Right - ts.Width < 0 Then
                                                Top += CurTop
                                                NextRight = MaximumSize.Width
                                                Right = NextRight
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If
                                CurTop = 0
                            End If
                            IsFirst = True
                        End If
                        SelectObject(hdc, oldFont)
                        g.ReleaseHdc(hdc)
                        g.Dispose()
                        NewText.Text = theText.Substring(0, nChar)
                        theText = theText.Substring(nChar)
                        Texts(Count)(SubCount).Add(NewText)
                        Me.Controls.Add(NewText)
                        NewText.Width = NewText.PreferredSize.Width
                        NewText.Top = Top + CurTop
                        If Right - NewText.PreferredSize.Width < 0 Then
                            CurTop += NewText.PreferredSize.Height
                            Right = MaximumSize.Width - NewText.PreferredSize.Width
                            NextRight = MaximumSize.Width
                            IsOverflow = True
                        Else
                            Right -= NewText.PreferredSize.Width
                        End If
                        NewText.Left = Right
                    End While
                    CurTop += Texts(Count)(SubCount)(Texts(Count)(SubCount).Count - 1).PreferredSize.Height
                End If
            Next
            If IsOverflow Then
                Top += CurTop
                CurTop = 0
            End If
            MaxRight = Math.Min(Right, MaxRight)
            If Count = _RenderArray.Items.Count - 1 Then
                Top += CurTop + Texts(Count)(_RenderArray.Items(Count).TextItems.Length - 1)(Texts(Count)(_RenderArray.Items(Count).TextItems.Length - 1).Count - 1).PreferredSize.Height
            End If
        Next
        Me.Size = New Size(MaximumSize.Width, Top) '- Math.Min(Right, MaxRight)
    End Sub

    Private Sub MultiLangRender_Load(sender As Object, e As EventArgs) Handles Me.Load
        RecalcLayout()
    End Sub
    Private Sub MultiLangRender_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        RecalcLayout()
    End Sub
End Class
