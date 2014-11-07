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
    Dim CurBounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo))) = Nothing
    Dim _ReEntry As Boolean = False
    Dim _RenderArray As Generic.List(Of IslamMetadata.RenderArray.RenderItem)
    Structure LayoutInfo
        Public Sub New(NewRect As Rectangle, NewNChar As Integer, NewBounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo))))
            Rect = NewRect
            nChar = NewNChar
            Bounds = NewBounds
        End Sub
        Dim Rect As Rectangle
        Dim nChar As Integer
        Dim Bounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo)))
    End Structure
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
        Dim Font As New PdfSharp.Drawing.XFont("Arial", 20, PdfSharp.Drawing.XFontStyle.Bold)
        Dim Bounds As New Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo)))
        'define page height and split
        'For Count = 0 To Pages.Count - 1
        Dim Page As New PdfSharp.Pdf.PdfPage
        GetLayout(_RenderArray, Page.Width.Point, Bounds, GetTextWidthFromPdf())
        Dim XGraphics As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(Page)
        'XGraphics.DrawString("test", Font, PdfSharp.Drawing.XBrushes.Black, New PdfSharp.Drawing.XPoint(x, y))
        'Next
        Doc.Save(Path)
    End Sub
    Delegate Function GetTextWidth(Str As String, MaxWidth As Double, ByRef s As Size) As Integer
    Private Shared Function GetTextWidthFromTextBox(NewText As TextBox, hdc As IntPtr) As GetTextWidth
        Dim ret As IntPtr = SendMessage(NewText.Handle, EM_GETMARGINS, IntPtr.Zero, IntPtr.Zero)
        Dim WidthOffset As Integer = (ret.ToInt32() And &HFFFF) + (ret.ToInt32() << 16) + GetSystemMetrics(SM_CXBORDER) * 2 + NewText.Margin.Left + NewText.Margin.Right
        Return Function(Str As String, MaxWidth As Double, ByRef s As Size)
                   Dim nChar As Integer
                   GetTextExtentExPoint(hdc, Str, Str.Length, CInt(MaxWidth) - WidthOffset, nChar, Nothing, s)
                   s.Width += WidthOffset
                   NewText.Text = Str
                   s.Height = NewText.PreferredSize.Height
                   Return nChar
               End Function
    End Function
    Private Shared Function GetTextWidthFromPdf() As GetTextWidth
        Return Function(Str As String, MaxWidth As Double, ByRef s As Size)
                   Return 0
               End Function
    End Function
    Private Shared Function GetLayout(CurRenderArray As List(Of IslamMetadata.RenderArray.RenderItem), _Width As Double, ByRef Bounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo))), WidthFunc As GetTextWidth) As Size
        Dim MaxRight As Double = _Width
        Dim Top As Integer = 0
        Dim NextRight As Double = _Width
        Dim LastCurTop As Integer = 0
        Dim LastRight As Double = _Width
        For Count As Integer = 0 To CurRenderArray.Count - 1
            Dim IsOverflow As Boolean = False
            Dim MaxWidth As Double = 0
            Dim Right As Double = NextRight
            Dim CurTop As Integer = 0
            Bounds.Add(New Generic.List(Of Generic.List(Of LayoutInfo)))
            For SubCount As Integer = 0 To CurRenderArray(Count).TextItems.Length - 1
                Bounds(Count).Add(New Generic.List(Of LayoutInfo))
                Dim s As Drawing.Size
                If CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eNested Then
                    Dim SubBounds As New Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo)))
                    s = MultiLangRender.GetLayout(CType(CurRenderArray(Count).TextItems(SubCount).Text, List(Of IslamMetadata.RenderArray.RenderItem)), _Width, SubBounds, WidthFunc)
                    Right = NextRight
                    If s.Width > NextRight Then
                        NextRight = _Width
                        IsOverflow = True
                    End If
                    Bounds(Count)(SubCount).Add(New LayoutInfo(New Rectangle(CInt(Right), Top + CurTop, s.Width, s.Height), 0, SubBounds))
                    MaxWidth = Math.Max(MaxWidth, s.Width)
                ElseIf CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eArabic Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eLTR Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eRTL Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eTransliteration Then
                    Dim theText As String = CStr(CurRenderArray(Count).TextItems(SubCount).Text)
                    While theText <> String.Empty
                        Dim nChar As Integer
                        nChar = WidthFunc(theText, _Width, s)
                        'break up string on previous word boundary unless beginning of string
                        If nChar = 0 Then
                            nChar = theText.Length 'If no room for even a letter than just use placeholder
                        ElseIf nChar <> theText.Length Then
                            Dim idx As Integer = Array.FindLastIndex(theText.ToCharArray(), nChar - 1, nChar, Function(ch As Char) Char.IsWhiteSpace(ch))
                            If idx <> -1 Then nChar = idx + 1
                        End If
                        If theText.Substring(nChar) <> String.Empty Then
                            WidthFunc(theText.Substring(0, nChar), _Width, s)
                        End If
                        theText = theText.Substring(nChar)
                        Right = NextRight
                        If theText <> String.Empty Or s.Width > NextRight Then
                            NextRight = _Width
                            IsOverflow = True
                        End If
                        Bounds(Count)(SubCount).Add(New LayoutInfo(New Rectangle(CInt(Right), Top + CurTop, s.Width, s.Height), nChar, Nothing))
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
                        Bounds(Count)(SubCount)(NextCount) = New LayoutInfo(New Rectangle(CInt(MaxWidth / 2 - Bounds(Count)(SubCount)(NextCount).Rect.Width / 2), Bounds(Count)(SubCount)(NextCount).Rect.Top + If(IsOverflow, LastCurTop, 0), Bounds(Count)(SubCount)(NextCount).Rect.Width, Bounds(Count)(SubCount)(NextCount).Rect.Height), Bounds(Count)(SubCount)(NextCount).nChar, Bounds(Count)(SubCount)(NextCount).Bounds)
                    Else
                        Bounds(Count)(SubCount)(NextCount) = New LayoutInfo(New Rectangle(If(IsOverflow, CInt(_Width), Bounds(Count)(SubCount)(NextCount).Rect.Left) - CInt(MaxWidth - (MaxWidth / 2 - Bounds(Count)(SubCount)(NextCount).Rect.Width / 2)), Bounds(Count)(SubCount)(NextCount).Rect.Top + If(IsOverflow, LastCurTop, 0), Bounds(Count)(SubCount)(NextCount).Rect.Width, Bounds(Count)(SubCount)(NextCount).Rect.Height), Bounds(Count)(SubCount)(NextCount).nChar, Bounds(Count)(SubCount)(NextCount).Bounds)
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
                    Bounds(Count)(SubCount)(NextCount) = New LayoutInfo(New Rectangle(Bounds(Count)(SubCount)(NextCount).Rect.Left - CInt(_Width - MaxRight), Bounds(Count)(SubCount)(NextCount).Rect.Top, Bounds(Count)(SubCount)(NextCount).Rect.Width, Bounds(Count)(SubCount)(NextCount).Rect.Height), Bounds(Count)(SubCount)(NextCount).nChar, Bounds(Count)(SubCount)(NextCount).Bounds)
                Next
            Next
        Next
        Return New Size(CInt(MaxRight), Top)
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
            Me.Size = GetLayout(_RenderArray, Me.Parent.Width, _Bounds, GetTextWidthFromTextBox(CalcText, hdc))
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
                    Renderer.Bounds = _Bounds(Count)(SubCount)(0).Rect
                ElseIf _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eArabic Or _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eLTR Or _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eRTL Or _RenderArray(Count).TextItems(SubCount).DisplayClass = IslamMetadata.RenderArray.RenderDisplayClass.eTransliteration Then
                    Dim theText As String = CStr(_RenderArray(Count).TextItems(SubCount).Text)
                    For NextCount As Integer = 0 To _Bounds(Count)(SubCount).Count - 1
                        Dim NewText As New TextBox
                        NewText.RightToLeft = If(_RenderArray(Count).TextItems(SubCount).DisplayClass <> IslamMetadata.RenderArray.RenderDisplayClass.eLTR And _RenderArray(Count).TextItems(SubCount).DisplayClass <> IslamMetadata.RenderArray.RenderDisplayClass.eTransliteration, Windows.Forms.RightToLeft.Yes, Windows.Forms.RightToLeft.No)
                        NewText.Font = Font
                        NewText.Text = theText.Substring(0, _Bounds(Count)(SubCount)(NextCount).nChar)
                        theText = theText.Substring(_Bounds(Count)(SubCount)(NextCount).nChar)
                        Me.Controls.Add(NewText)
                        NewText.Bounds = _Bounds(Count)(SubCount)(NextCount).Rect
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
