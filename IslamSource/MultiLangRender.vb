Public Class NativeMethods
    <Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="GetSystemMetrics")> _
    Friend Shared Function GetSystemMetrics(ByVal nIndex As Integer) As Integer
    End Function
    <Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SendMessage")> _
    Friend Shared Function SendMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("gdi32.dll", EntryPoint:="SelectObject")> _
    Friend Shared Function SelectObject(ByVal hdc As IntPtr, ByVal hObject As IntPtr) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("gdi32.dll", EntryPoint:="GetTextExtentExPoint", SetLastError:=True, CharSet:=Runtime.InteropServices.CharSet.Unicode)> _
    Friend Shared Function GetTextExtentExPoint(ByVal hdc As IntPtr, <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.LPWStr)> ByVal lpszStr As String, ByVal cchString As Integer, ByVal nMaxExtent As Integer, ByRef lpnFit As Integer, ByVal alpDx As Integer(), ByRef lpSize As Size) As Boolean
    End Function
    <Runtime.InteropServices.DllImport("gdi32.dll", EntryPoint:="SetTextAlign")> _
    Friend Shared Function SetTextAlign(ByVal hdc As IntPtr, ByVal fMode As UInteger) As UInteger
    End Function
End Class
Public Class MultiLangRender
    Const TA_RTLREADING As UInteger = 256
    Const SM_CXBORDER As Integer = 5
    Const SM_CYBORDER As Integer = 6
    Const EM_GETMARGINS As UInteger = &HD4
    Dim CurBounds As Generic.List(Of Generic.List(Of Generic.List(Of HostPageUtility.RenderArrayWeb.LayoutInfo))) = Nothing
    Dim _ReEntry As Boolean = False
    Dim _RenderArray As Generic.List(Of XMLRender.RenderArray.RenderItem)
    Public Property RenderArray As List(Of XMLRender.RenderArray.RenderItem)
        Get
            Return _RenderArray
        End Get
        Set(value As List(Of XMLRender.RenderArray.RenderItem))
            _RenderArray = value
        End Set
    End Property

    Private Shared Function GetTextWidthFromTextBox(NewText As TextBox, hdc As IntPtr) As HostPageUtility.RenderArrayWeb.GetTextWidth
        Dim ret As IntPtr = NativeMethods.SendMessage(NewText.Handle, EM_GETMARGINS, IntPtr.Zero, IntPtr.Zero)
        Dim WidthOffset As Integer = (ret.ToInt32() And &HFFFF) + (ret.ToInt32() << 16) + NativeMethods.GetSystemMetrics(SM_CXBORDER) * 2 + NewText.Margin.Left + NewText.Margin.Right
        Return Function(Str As String, FontName As String, MaxWidth As Single, IsRTL As Boolean, ByRef s As SizeF, ByRef Baseline As Single)
                   Dim nChar As Integer
                   Dim GetSize As Size
                   Baseline = NewText.Font.FontFamily.GetCellAscent(NewText.Font.Style)
                   NativeMethods.SetTextAlign(hdc, If(IsRTL, TA_RTLREADING, CUInt(0)))
                   NativeMethods.GetTextExtentExPoint(hdc, Str, Str.Length, CInt(MaxWidth) - WidthOffset, nChar, Nothing, GetSize)
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
        Dim _Bounds As Generic.List(Of Generic.List(Of Generic.List(Of HostPageUtility.RenderArrayWeb.LayoutInfo))) = CurBounds
        If _Bounds Is Nothing Then
            _Bounds = New Generic.List(Of Generic.List(Of Generic.List(Of HostPageUtility.RenderArrayWeb.LayoutInfo)))
            Me.RightToLeft = Windows.Forms.RightToLeft.Yes
            Me.MaximumSize = New Size(Me.Parent.Width, Me.Height)
            Dim g As Graphics = CreateGraphics()
            Dim hdc As IntPtr = g.GetHdc()
            Dim oldFont As IntPtr = NativeMethods.SelectObject(hdc, Font.ToHfont())
            Dim CalcText As New TextBox
            CalcText.Font = Font
            Me.Size = HostPageUtility.RenderArrayWeb.GetLayout(_RenderArray, CSng(Me.Parent.Width), _Bounds, GetTextWidthFromTextBox(CalcText, hdc)).ToSize()
            NativeMethods.SelectObject(hdc, oldFont)
            g.ReleaseHdc(hdc)
            g.Dispose()
        End If
        For Count As Integer = 0 To _RenderArray.Count - 1
            For SubCount As Integer = 0 To _RenderArray(Count).TextItems.Length - 1
                If _RenderArray(Count).TextItems(SubCount).DisplayClass = XMLRender.RenderArray.RenderDisplayClass.eNested Then
                    Dim Renderer As New MultiLangRender
                    Renderer.CurBounds = _Bounds(Count)(SubCount)(0).Bounds
                    Renderer.RenderArray = CType(_RenderArray(Count).TextItems(SubCount).Text, List(Of XMLRender.RenderArray.RenderItem))
                    Me.Controls.Add(Renderer)
                    Renderer.Bounds = Rectangle.Round(_Bounds(Count)(SubCount)(0).Rect)
                ElseIf _RenderArray(Count).TextItems(SubCount).DisplayClass = XMLRender.RenderArray.RenderDisplayClass.eArabic Or _RenderArray(Count).TextItems(SubCount).DisplayClass = XMLRender.RenderArray.RenderDisplayClass.eLTR Or _RenderArray(Count).TextItems(SubCount).DisplayClass = XMLRender.RenderArray.RenderDisplayClass.eRTL Or _RenderArray(Count).TextItems(SubCount).DisplayClass = XMLRender.RenderArray.RenderDisplayClass.eTransliteration Then
                    Dim theText As String = CStr(_RenderArray(Count).TextItems(SubCount).Text)
                    For NextCount As Integer = 0 To _Bounds(Count)(SubCount).Count - 1
                        Dim NewText As New TextBox
                        NewText.RightToLeft = If(_RenderArray(Count).TextItems(SubCount).DisplayClass <> XMLRender.RenderArray.RenderDisplayClass.eLTR And _RenderArray(Count).TextItems(SubCount).DisplayClass <> XMLRender.RenderArray.RenderDisplayClass.eTransliteration, Windows.Forms.RightToLeft.Yes, Windows.Forms.RightToLeft.No)
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
