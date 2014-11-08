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
    <Runtime.InteropServices.DllImport("gdi32.dll", EntryPoint:="SetTextAlign")> _
    Private Shared Function SetTextAlign(ByVal hdc As IntPtr, ByVal fMode As UInteger) As UInteger
    End Function
    Const TA_RTLREADING As UInteger = 256
    Const SM_CXBORDER As Integer = 5
    Const SM_CYBORDER As Integer = 6
    Const EM_GETMARGINS As UInteger = &HD4
    Dim CurBounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo))) = Nothing
    Dim _ReEntry As Boolean = False
    Dim _RenderArray As Generic.List(Of IslamMetadata.RenderArray.RenderItem)
    Public Property RenderArray As List(Of IslamMetadata.RenderArray.RenderItem)
        Get
            Return _RenderArray
        End Get
        Set(value As List(Of IslamMetadata.RenderArray.RenderItem))
            _RenderArray = value
        End Set
    End Property
    Structure LayoutInfo
        Public Sub New(NewRect As RectangleF, NewNChar As Integer, NewBounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo))))
            Rect = NewRect
            nChar = NewNChar
            Bounds = NewBounds
        End Sub
        Dim Rect As RectangleF
        Dim nChar As Integer
        Dim Bounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo)))
    End Structure
    Public Shared Sub DoRenderPdf(Doc As iTextSharp.text.Document, Writer As iTextSharp.text.pdf.PdfWriter, Font As iTextSharp.text.Font, CurRenderArray As List(Of IslamMetadata.RenderArray.RenderItem), _Bounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo))), ByRef PageOffset As PointF, BaseOffset As PointF)
        For Count As Integer = 0 To CurRenderArray.Count - 1
            For SubCount As Integer = 0 To CurRenderArray(Count).TextItems.Length - 1
                If CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eNested Then
                    DoRenderPdf(Doc, Writer, Font, CType(CurRenderArray(Count).TextItems(SubCount).Text, List(Of IslamMetadata.RenderArray.RenderItem)), _Bounds(Count)(SubCount)(0).Bounds, PageOffset, New PointF(_Bounds(Count)(SubCount)(0).Rect.Location.X, _Bounds(Count)(SubCount)(0).Rect.Location.Y))
                ElseIf CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eArabic Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eLTR Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eRTL Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eTransliteration Then
                    Dim theText As String = CStr(CurRenderArray(Count).TextItems(SubCount).Text)
                    For NextCount As Integer = 0 To _Bounds(Count)(SubCount).Count - 1
                        If _Bounds(Count)(SubCount)(NextCount).Rect.Top + PageOffset.Y + BaseOffset.Y > Doc.PageSize.Height - Doc.BottomMargin - Doc.TopMargin Then
                            Doc.NewPage()
                            PageOffset.Y = -_Bounds(Count)(SubCount)(NextCount).Rect.Top - BaseOffset.Y
                            Exit For
                        End If
                    Next
                    For NextCount As Integer = 0 To _Bounds(Count)(SubCount).Count - 1
                        Dim Rect As RectangleF = _Bounds(Count)(SubCount)(NextCount).Rect
                        Rect.Offset(BaseOffset)
                        Rect.Offset(PageOffset)
                        Dim ct As New iTextSharp.text.pdf.ColumnText(Writer.DirectContent)
                        If CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eArabic Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eRTL Then
                            ct.RunDirection = iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL
                            ct.ArabicOptions = iTextSharp.text.pdf.ColumnText.AR_LIG
                            ct.UseAscender = False
                        Else
                            ct.RunDirection = iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_LTR
                        End If
                        ct.SetSimpleColumn(Rect.Left + Doc.LeftMargin, Doc.PageSize.Height - Doc.BottomMargin - Doc.TopMargin - Rect.Bottom, Rect.Right + 1 + Doc.LeftMargin, Doc.PageSize.Height - Doc.BottomMargin - Doc.TopMargin - Rect.Top + 1, Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_LEADING, Font.Size), iTextSharp.text.Element.ALIGN_CENTER)
                        ct.AddText(New iTextSharp.text.Chunk(theText.Substring(0, _Bounds(Count)(SubCount)(NextCount).nChar), Font))
                        ct.Go()
                        theText = theText.Substring(_Bounds(Count)(SubCount)(NextCount).nChar)
                    Next
                End If
            Next
        Next
    End Sub
    Public Shared Function GetFontPath(Index As Integer) As String
        'Return IslamMetadata.Utility.GetFilePath("files\" + "Scheherazade-R.ttf")
        Dim Fonts As String() = {"times.ttf", "me_quran.ttf", "Scheherazade.ttf", "PDMS_Saleem.ttf", "KFC_naskh.otf", "trado.ttf", "arabtype.ttf", "majalla.ttf", "msuighur.ttf", "ARIALUNI.ttf"}
        Return If(Index < 1 Or Index > 4, IO.Path.Combine(Environment.GetEnvironmentVariable("windir"), "Fonts\" + Fonts(Index)), IslamMetadata.Utility.GetFilePath("files\" + Fonts(Index)))
    End Function
    Public Shared Sub OutputPdf(Path As String, CurRenderArray As List(Of IslamMetadata.RenderArray.RenderItem))
        Dim Doc As New iTextSharp.text.Document
        Dim Writer As iTextSharp.text.pdf.PdfWriter = iTextSharp.text.pdf.PdfWriter.GetInstance(Doc, New IO.FileStream(Path, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.None))
        Doc.Open()
        Doc.NewPage()
        Dim BaseFont As iTextSharp.text.pdf.BaseFont = iTextSharp.text.pdf.BaseFont.CreateFont(GetFontPath(5), iTextSharp.text.pdf.BaseFont.IDENTITY_H, iTextSharp.text.pdf.BaseFont.NOT_EMBEDDED)
        Dim Font As New iTextSharp.text.Font(BaseFont, 20, iTextSharp.text.Font.BOLD)
        Dim _Bounds As New Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo)))
        'divide into pages by heights
        GetLayout(CurRenderArray, Doc.PageSize.Width - Doc.LeftMargin - Doc.RightMargin, _Bounds, GetTextWidthFromPdf(Font, Writer.DirectContent))
        Dim PageOffset As New PointF(0, 0)
        DoRenderPdf(Doc, Writer, Font, CurRenderArray, _Bounds, PageOffset, New PointF(0, 0))
        Doc.Close()
    End Sub
    Delegate Function GetTextWidth(Str As String, MaxWidth As Single, IsRTL As Boolean, ByRef s As SizeF) As Integer
    Private Shared Function GetTextWidthFromPdf(Font As iTextSharp.text.Font, Content As iTextSharp.text.pdf.PdfContentByte) As GetTextWidth
        Return Function(Str As String, MaxWidth As Single, IsRTL As Boolean, ByRef s As SizeF)
                   Font.BaseFont.CorrectArabicAdvance()
                   s.Height = 8 * Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_LEADING, Font.Size) + Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_ASCENT, Font.Size) - Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_DESCENT, Font.Size)
                   s.Width = iTextSharp.text.pdf.ColumnText.GetWidth(New iTextSharp.text.Phrase(New iTextSharp.text.Chunk(Str, Font)), If(IsRTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_LTR), iTextSharp.text.pdf.ColumnText.AR_COMPOSEDTASHKEEL)
                   Dim Len As Integer = Str.Length
                   Dim Search As Integer = Len
                   'binary search the maximum characters
                   If s.Width > MaxWidth Then
                       While Search <> 1
                           Search = Search \ 2
                           If s.Width > MaxWidth Then
                               Len -= Search
                           Else
                               Len += Search
                           End If
                           s.Width = iTextSharp.text.pdf.ColumnText.GetWidth(New iTextSharp.text.Phrase(Str.Substring(0, Len), Font), If(IsRTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_LTR), 0)
                       End While
                       If s.Width > MaxWidth Then
                           Len -= 1 'factor towards fitting not overflowing
                           s.Width = iTextSharp.text.pdf.ColumnText.GetWidth(New iTextSharp.text.Phrase(Str.Substring(0, Len), Font), If(IsRTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_LTR), 0)
                       End If
                   End If
                   Return Len
               End Function
    End Function
    Private Shared Function GetLayout(CurRenderArray As List(Of IslamMetadata.RenderArray.RenderItem), _Width As Single, ByRef Bounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo))), WidthFunc As GetTextWidth) As SizeF
        Dim MaxRight As Single = _Width
        Dim Top As Single = 0
        Dim NextRight As Single = _Width
        Dim LastCurTop As Single = 0
        Dim LastRight As Single = _Width
        For Count As Integer = 0 To CurRenderArray.Count - 1
            Dim IsOverflow As Boolean = False
            Dim MaxWidth As Single = 0
            Dim Right As Single = NextRight
            Dim CurTop As Single = 0
            Bounds.Add(New Generic.List(Of Generic.List(Of LayoutInfo)))
            For SubCount As Integer = 0 To CurRenderArray(Count).TextItems.Length - 1
                Bounds(Count).Add(New Generic.List(Of LayoutInfo))
                Dim s As Drawing.SizeF
                If CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eNested Then
                    Dim SubBounds As New Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo)))
                    s = MultiLangRender.GetLayout(CType(CurRenderArray(Count).TextItems(SubCount).Text, List(Of IslamMetadata.RenderArray.RenderItem)), _Width, SubBounds, WidthFunc)
                    Right = NextRight
                    If s.Width > NextRight Then
                        NextRight = _Width
                        IsOverflow = True
                    End If
                    Bounds(Count)(SubCount).Add(New LayoutInfo(New RectangleF(Right, Top + CurTop, s.Width, s.Height), 0, SubBounds))
                    MaxWidth = Math.Max(MaxWidth, s.Width)
                ElseIf CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eArabic Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eLTR Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eRTL Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eTransliteration Then
                    Dim theText As String = CStr(CurRenderArray(Count).TextItems(SubCount).Text)
                    While theText <> String.Empty
                        Dim nChar As Integer
                        nChar = WidthFunc(theText, _Width, CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eArabic Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eRTL, s)
                        'break up string on previous word boundary unless beginning of string
                        If nChar = 0 Then
                            nChar = theText.Length 'If no room for even a letter than just use placeholder
                        ElseIf nChar <> theText.Length Then
                            Dim idx As Integer = Array.FindLastIndex(theText.ToCharArray(), nChar - 1, nChar, Function(ch As Char) Char.IsWhiteSpace(ch))
                            If idx <> -1 Then nChar = idx + 1
                        End If
                        If theText.Substring(nChar) <> String.Empty Then
                            WidthFunc(theText.Substring(0, nChar), _Width, CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eArabic Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eRTL, s)
                        End If
                        theText = theText.Substring(nChar)
                        Right = NextRight
                        If theText <> String.Empty Or s.Width > NextRight Then
                            NextRight = _Width
                            IsOverflow = True
                        End If
                        Bounds(Count)(SubCount).Add(New LayoutInfo(New RectangleF(Right, Top + CurTop, s.Width, s.Height), nChar, Nothing))
                        If theText <> String.Empty Then CurTop += s.Height
                        MaxWidth = Math.Max(MaxWidth, s.Width)
                    End While
                End If
                If Bounds(Count)(SubCount).Count <> 0 Then
                    CurTop += s.Height
                End If
            Next
            'centering must come after maximum width is calculated
            For SubCount = 0 To Bounds(Count).Count - 1
                For NextCount = 0 To Bounds(Count)(SubCount).Count - 1
                    If NextCount <> Bounds(Count)(SubCount).Count - 1 Then
                        Bounds(Count)(SubCount)(NextCount) = New LayoutInfo(New RectangleF(MaxWidth / 2 - Bounds(Count)(SubCount)(NextCount).Rect.Width / 2, Bounds(Count)(SubCount)(NextCount).Rect.Top + If(IsOverflow, LastCurTop, 0), Bounds(Count)(SubCount)(NextCount).Rect.Width, Bounds(Count)(SubCount)(NextCount).Rect.Height), Bounds(Count)(SubCount)(NextCount).nChar, Bounds(Count)(SubCount)(NextCount).Bounds)
                    Else
                        Bounds(Count)(SubCount)(NextCount) = New LayoutInfo(New RectangleF(If(IsOverflow, _Width, Bounds(Count)(SubCount)(NextCount).Rect.Left) - CInt(MaxWidth - (MaxWidth / 2 - Bounds(Count)(SubCount)(NextCount).Rect.Width / 2)), Bounds(Count)(SubCount)(NextCount).Rect.Top + If(IsOverflow, LastCurTop, 0), Bounds(Count)(SubCount)(NextCount).Rect.Width, Bounds(Count)(SubCount)(NextCount).Rect.Height), Bounds(Count)(SubCount)(NextCount).nChar, Bounds(Count)(SubCount)(NextCount).Bounds)
                    End If
                Next
            Next
            If IsOverflow Then
                Top += CurTop + LastCurTop
                CurTop = 0
                Right = NextRight
            Else
                NextRight -= MaxWidth
            End If
            LastCurTop = CurTop
            LastRight = NextRight
            If Count = CurRenderArray.Count - 1 Then
                Top += CurTop + Bounds(Count)(CurRenderArray(Count).TextItems.Length - 1)(Bounds(Count)(CurRenderArray(Count).TextItems.Length - 1).Count - 1).Rect.Height
            End If
            MaxRight = Math.Min(Math.Max(Math.Max(Right, _Width - MaxWidth), Right + MaxWidth), MaxRight)
        Next
        For Count = 0 To Bounds.Count - 1
            For SubCount = 0 To Bounds(Count).Count - 1
                For NextCount = 0 To Bounds(Count)(SubCount).Count - 1
                    'overall centering can be done here though must calculate an overall line width
                    Bounds(Count)(SubCount)(NextCount) = New LayoutInfo(New RectangleF(Bounds(Count)(SubCount)(NextCount).Rect.Left - CInt(_Width - MaxRight), Bounds(Count)(SubCount)(NextCount).Rect.Top, Bounds(Count)(SubCount)(NextCount).Rect.Width, Bounds(Count)(SubCount)(NextCount).Rect.Height), Bounds(Count)(SubCount)(NextCount).nChar, Bounds(Count)(SubCount)(NextCount).Bounds)
                Next
            Next
        Next
        Return New SizeF(MaxRight, Top)
    End Function

    Private Shared Function GetTextWidthFromTextBox(NewText As TextBox, hdc As IntPtr) As GetTextWidth
        Dim ret As IntPtr = SendMessage(NewText.Handle, EM_GETMARGINS, IntPtr.Zero, IntPtr.Zero)
        Dim WidthOffset As Integer = (ret.ToInt32() And &HFFFF) + (ret.ToInt32() << 16) + GetSystemMetrics(SM_CXBORDER) * 2 + NewText.Margin.Left + NewText.Margin.Right
        Return Function(Str As String, MaxWidth As Single, IsRTL As Boolean, ByRef s As SizeF)
                   Dim nChar As Integer
                   Dim GetSize As Size
                   SetTextAlign(hdc, If(IsRTL, TA_RTLREADING, CUInt(0)))
                   GetTextExtentExPoint(hdc, Str, Str.Length, CInt(MaxWidth) - WidthOffset, nChar, Nothing, GetSize)
                   s = New SizeF(GetSize.Width, GetSize.Height)
                   s.Width += WidthOffset
                   NewText.Text = Str
                   s.Height = NewText.PreferredSize.Height
                   Return nChar
               End Function
    End Function
    Private Sub RecalcLayout()
        Me.Controls.Clear()
        If _RenderArray Is Nothing Or Me.Parent Is Nothing OrElse Me.Parent.Width = 0 Then Return
        Dim _Bounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo))) = CurBounds
        If _Bounds Is Nothing Then
            _Bounds = New Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo)))
            Me.RightToLeft = Windows.Forms.RightToLeft.Yes
            Me.MaximumSize = New Size(Me.Parent.Width, Me.Height)
            Dim g As Graphics = CreateGraphics()
            Dim hdc As IntPtr = g.GetHdc()
            Dim oldFont As IntPtr = SelectObject(hdc, Font.ToHfont())
            Dim CalcText As New TextBox
            CalcText.Font = Font
            Me.Size = GetLayout(_RenderArray, CSng(Me.Parent.Width), _Bounds, GetTextWidthFromTextBox(CalcText, hdc)).ToSize()
            SelectObject(hdc, oldFont)
            g.ReleaseHdc(hdc)
            g.Dispose()
        End If
        For Count As Integer = 0 To _RenderArray.Count - 1
            For SubCount As Integer = 0 To _RenderArray(Count).TextItems.Length - 1
                If _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eNested Then
                    Dim Renderer As New MultiLangRender
                    Renderer.CurBounds = _Bounds(Count)(SubCount)(0).Bounds
                    Renderer.RenderArray = CType(_RenderArray(Count).TextItems(SubCount).Text, List(Of IslamMetadata.RenderArray.RenderItem))
                    Me.Controls.Add(Renderer)
                    Renderer.Bounds = Rectangle.Round(_Bounds(Count)(SubCount)(0).Rect)
                ElseIf _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eArabic Or _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eLTR Or _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eRTL Or _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eTransliteration Then
                    Dim theText As String = CStr(_RenderArray(Count).TextItems(SubCount).Text)
                    For NextCount As Integer = 0 To _Bounds(Count)(SubCount).Count - 1
                        Dim NewText As New TextBox
                        NewText.RightToLeft = If(_RenderArray(Count).TextItems(SubCount).DisplayClass <> IslamMetadata.RenderArray.RenderDisplayClass.eLTR And _RenderArray(Count).TextItems(SubCount).DisplayClass <> IslamMetadata.RenderArray.RenderDisplayClass.eTransliteration, Windows.Forms.RightToLeft.Yes, Windows.Forms.RightToLeft.No)
                        NewText.Font = Font
                        NewText.Text = theText.Substring(0, _Bounds(Count)(SubCount)(NextCount).nChar)
                        theText = theText.Substring(_Bounds(Count)(SubCount)(NextCount).nChar)
                        Me.Controls.Add(NewText)
                        NewText.Bounds = Rectangle.Round(_Bounds(Count)(SubCount)(NextCount).Rect)
                    Next
                End If
            Next
        Next
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
