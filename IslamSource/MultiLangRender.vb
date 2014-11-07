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
    Dim _ReEntry As Boolean = False
    Dim _RenderArray As Generic.List(Of IslamMetadata.RenderArray.RenderItem)
    Dim MaxWidths As New Generic.List(Of Integer)
    Dim Texts As New Generic.List(Of Generic.List(Of Generic.List(Of Object)))
    Public Property RenderArray As List(Of IslamMetadata.RenderArray.RenderItem)
        Get
            Return _RenderArray
        End Get
        Set(value As List(Of IslamMetadata.RenderArray.RenderItem))
            _RenderArray = value
        End Set
    End Property
    Public Sub OutputPdf(Path As String)
        Dim Doc As New PdfSharp.Pdf.PdfDocument
        'define page height and split
        'For Count = 0 To Pages.Count -1
        Dim Page As New PdfSharp.Pdf.PdfPage
        Dim XGraphics As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(Page)
        Dim Font As New PdfSharp.Drawing.XFont("Arial", 20, PdfSharp.Drawing.XFontStyle.Bold)
        'XGraphics.DrawString("test", Font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XPoint(x, y))
        'Next
        Doc.Save(Path)
    End Sub
    Private Sub RecalcLayout()
        Me.Controls.Clear()
        Texts.Clear()
        MaxWidths.Clear()
        If _RenderArray Is Nothing Or Me.Parent Is Nothing OrElse Me.Parent.Width = 0 Then Return
        Me.RightToLeft = Windows.Forms.RightToLeft.Yes

        Dim g As Graphics = CreateGraphics()
        Dim hdc As IntPtr = g.GetHdc()
        Dim oldFont As IntPtr = SelectObject(hdc, Font.ToHfont())
        For Count As Integer = 0 To _RenderArray.Count - 1
            MaxWidths.Add(0)
            Texts.Add(New Generic.List(Of Generic.List(Of Object)))
            For SubCount As Integer = 0 To _RenderArray(Count).TextItems.Length - 1
                Texts(Count).Add(New Generic.List(Of Object))
                If _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eNested Then
                    Dim Renderer As New MultiLangRender
                    Renderer.RenderArray = CType(_RenderArray(Count).TextItems(SubCount).Text, List(Of IslamMetadata.RenderArray.RenderItem))
                    Texts(Count)(SubCount).Add(Renderer)
                    MaxWidths(Count) = Math.Max(MaxWidths(Count), Renderer.Width)
                ElseIf _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eArabic Or _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eLTR Or _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eRTL Or _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eTransliteration Then
                    Dim theText As String = CStr(_RenderArray(Count).TextItems(SubCount).Text)
                    While theText <> String.Empty
                        Dim NewText As New TextBox
                        NewText.RightToLeft = If(_RenderArray(Count).TextItems(SubCount).DisplayClass <> IslamMetadata.RenderArray.RenderDisplayClass.eLTR And _RenderArray(Count).TextItems(SubCount).DisplayClass <> IslamMetadata.RenderArray.RenderDisplayClass.eTransliteration, Windows.Forms.RightToLeft.Yes, Windows.Forms.RightToLeft.No)
                        NewText.Font = Font
                        Dim s As Drawing.Size
                        Dim nChar As Integer
                        Dim ret As IntPtr = SendMessage(NewText.Handle, EM_GETMARGINS, IntPtr.Zero, IntPtr.Zero)
                        GetTextExtentExPoint(hdc, theText, theText.Length, Me.Parent.Width - (ret.ToInt32() And &HFFFF) - (ret.ToInt32() << 16) - GetSystemMetrics(SM_CXBORDER) * 2 - NewText.Margin.Left - NewText.Margin.Right, nChar, Nothing, s)
                        'break up string on previous word boundary unless beginning of string
                        If nChar = 0 Then
                            nChar = theText.Length 'If no room for even a letter than just use placeholder
                        ElseIf nChar <> theText.Length Then
                            Dim idx As Integer = Array.FindLastIndex(theText.ToCharArray(), nChar - 1, nChar, Function(ch As Char) Char.IsWhiteSpace(ch))
                            If idx <> -1 Then nChar = idx + 1
                        End If
                        NewText.Text = theText.Substring(0, nChar)
                        theText = theText.Substring(nChar)
                        If theText <> String.Empty Then
                            GetTextExtentExPoint(hdc, NewText.Text, NewText.Text.Length, Me.Parent.Width - (ret.ToInt32() And &HFFFF) - (ret.ToInt32() << 16) - GetSystemMetrics(SM_CXBORDER) * 2 - NewText.Margin.Left - NewText.Margin.Right, nChar, Nothing, s)
                        End If
                        Texts(Count)(SubCount).Add(NewText)
                        NewText.Width = s.Width + (ret.ToInt32() And &HFFFF) + (ret.ToInt32() << 16) + GetSystemMetrics(SM_CXBORDER) * 2 + NewText.Margin.Left + NewText.Margin.Right
                        MaxWidths(Count) = Math.Max(MaxWidths(Count), NewText.Width)
                    End While
                End If
            Next
        Next
        SelectObject(hdc, oldFont)
        g.ReleaseHdc(hdc)
        g.Dispose()

        Dim CurTop As Integer = 0
        Dim Top As Integer = 0
        Dim Right As Integer = Me.Parent.Width
        'Dim MaxRight As Integer = Right
        Dim NextRight As Integer = Right
        Me.MaximumSize = New Size(Me.Parent.Width, Me.Height)
        Me.Size = New Size(Me.Parent.Width, Me.Height) 'Preset width for child calculations
        For Count As Integer = 0 To _RenderArray.Count - 1
            Dim IsOverflow As Boolean = False
            If CurTop <> 0 Then
                If Right - MaxWidths(Count) < 0 Then
                    Top += CurTop
                    NextRight = Me.Parent.Width
                    Right = NextRight
                End If
                CurTop = 0
            End If
            For SubCount As Integer = 0 To _RenderArray(Count).TextItems.Length - 1
                For NextCount As Integer = 0 To Texts(Count)(SubCount).Count - 1
                    Right = NextRight
                    CType(Texts(Count)(SubCount)(NextCount), Control).Top = Top + CurTop
                    If NextCount <> Texts(Count)(SubCount).Count - 1 Then
                        CurTop += CType(Texts(Count)(SubCount)(NextCount), Control).PreferredSize.Height
                        Right = CInt(MaxWidths(Count) / 2 - CType(Texts(Count)(SubCount)(NextCount), Control).Width / 2)
                        NextRight = Me.Parent.Width
                        IsOverflow = True
                    Else
                        Right -= CInt(MaxWidths(Count) - (MaxWidths(Count) / 2 - CType(Texts(Count)(SubCount)(NextCount), Control).Width / 2))
                    End If
                    CType(Texts(Count)(SubCount)(NextCount), Control).Left = Right
                    Me.Controls.Add(CType(Texts(Count)(SubCount)(NextCount), Control))
                Next
                If Texts(Count)(SubCount).Count <> 0 Then
                    CurTop += CType(Texts(Count)(SubCount)(Texts(Count)(SubCount).Count - 1), Control).PreferredSize.Height
                End If
            Next
            If IsOverflow Then
                Top += CurTop
                CurTop = 0
                Right = NextRight
            Else
                NextRight -= MaxWidths(Count)
            End If
            'MaxRight = Math.Min(Right, MaxRight)
            If Count = _RenderArray.Count - 1 Then
                Top += CurTop + CType(Texts(Count)(_RenderArray(Count).TextItems.Length - 1)(Texts(Count)(_RenderArray(Count).TextItems.Length - 1).Count - 1), Control).PreferredSize.Height
            End If
        Next
        Me.Size = New Size(Me.Parent.Width, Top) '- Math.Min(Right, MaxRight)
    End Sub

    Private Sub MultiLangRender_Load(sender As Object, e As EventArgs) Handles Me.Load
        Font = New System.Drawing.Font(Font.FontFamily.Name, 22, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        If Not _ReEntry Then
            _ReEntry = True
            RecalcLayout()
            _ReEntry = False
        End If
    End Sub
    Private Sub MultiLangRender_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Not _ReEntry Then
            _ReEntry = True
            RecalcLayout()
            _ReEntry = False
        End If
    End Sub
End Class
