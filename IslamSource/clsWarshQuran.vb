Public Class clsWarshQuran
    Shared Arr1 As New Dictionary(Of String, Integer()) From {
{" ", New Integer() {&H20}},
{"llh", New Integer() {&H21}},
{"A", New Integer() {&H23, &H24, &H25}},
{">", New Integer() {&H26, &H27, &H28}},
{"<", New Integer() {&H29, &H2A}},
{"|", New Integer() {&H2C, &H2D}},
{"b", New Integer() {&H2F, &H30, &H31, &H32, &H33, &H35, &H36, &H37, &H39, &H3A, &H3B, &H3C, &H3D, &H3E}},
{"t", New Integer() {&H3F, &H40, &H41, &H42, &H43, &H45, &H46, &H47, &H48, &H49, &H4A, &H4B, &H4C, &H4D, &H4E}},
{"v", New Integer() {&H4F, &H52, &H53, &H56, &H57, &H59, &H5A, &H5B, &H5C, &H5D, &H5E}},
{"j", New Integer() {&H5F, &H60, &H62, &H64, &H65, &H66, &H67, &H6B, &H6C}},
{"H", New Integer() {&H6D, &H6E, &H6F, &H70, &H71, &H72, &H73, &H74, &H76, &H77, &H78, &H79}},
{"x", New Integer() {&H7A, &H7B, &H7D, &H7E, &HF081, &H201A, &H192, &H2020, &H2021, &H2C6}},
{"d", New Integer() {&H2030, &H160}},
{"*", New Integer() {&H2039, &H152}},
{"r", New Integer() {&HF08D, &HF08E, &HF08F, &HF090, &H2018, &H2019}},
{"z", New Integer() {&H201C, &H201D, &H2022, &H2013, &H2014}},
{"s", New Integer() {&H2122, &H203A, &H153, &HF09D, &HF09E, &HA0, &HA1, &HA3, &HA4, &HA6, &HA7, &HA8}},
{"$", New Integer() {&HA9, &HAB, &HAC, &H2212, &HB0, &HB1, &HB4, &HB5, &HB6, &HB7, &HB8}},
{"S", New Integer() {&HB9, &HBB, &HBC, &HBE, &HC0, &HC1, &HC2, &HC4, &HC7, &HC8, &HC9}},
{"D", New Integer() {&HCA, &HCC, &HCE, &HCF, &HD1, &HD2, &HD3, &HD5, &HD6, &HD8, &HD9, &HDA}},
{"T", New Integer() {&HDB, &HDC, &HDD, &HDE}},
{"Z", New Integer() {&HDF, &HE0, &HE1, &HE2}},
{"E", New Integer() {&HE3, &HE4, &HE5, &HE6, &HE7, &HE8, &HE9, &HEA, &HEB, &HEC, &HED}},
{"g", New Integer() {&HEE, &HEF, &HF0, &HF1, &HF3, &HF6, &HF0F7, &HF8}},
{"f", New Integer() {&HF9, &HF0FA, &HFB, &HF0FF}}
}
    Shared Arr2 As New Dictionary(Of String, Integer()) From {
    {" ", New Integer() {&H20}},
    {"f", New Integer() {&H2200, &H23, &H2203}},
    {"q", New Integer() {&H25, &H220B, &H29, &H2217, &H2B, &H2C, &H2212}},
    {"k", New Integer() {&H2E, &H30, &H31, &H32, &H33, &H34, &H35, &H36, &H37, &H38}},
    {"l", New Integer() {&H39, &H3A, &H3B, &H3C, &H3D, &H3E, &H3F, &H2245, &H391}},
    {"m", New Integer() {&H392, &H3A7, &H2206, &H395, &H393, &H397, &H399, &H3D1, &H39A, &H39B, &H39C, &H39D, &H39F, &H3A0, &H398}},
    {"n", New Integer() {&H3A1, &H3A3, &H3A4, &H3A5, &H3C2, &H2126, &H39E, &H3A8, &H396, &H2234, &H5D, &H22A5, &H5F, &H23AF, &H3B2, &H3C7, &H2205, &H239D, &HF0, &H239E, &H23A5, &H23A6}},
    {"h", New Integer() {&H3B4, &H3B5, &H3C6, &H3B3, &H3B7, &H3B9, &H3D5, &H3BA, &H3BB, &HB5, &H3BD}},
    {"p", New Integer() {&H3BF, &H3C0}},
    {"w", New Integer() {&H3B8, &H3C1}},
    {"&", New Integer() {&H3C3, &H3C4, &H3C5}},
    {"lA", New Integer() {&H3C9, &H3BE}},
    {"l>", New Integer() {&H3C8, &H7B, &H7C, &H7D, &H223C}},
    {"y", New Integer() {&H192, &H201E, &H2020, &H2021, &H2C6, &H2030, &H160, &H2039, &HF08D, &HF08E, &HF08F, &HF090, &H23A7, &H23AA, &H23AB}},
    {"Y", New Integer() {&H2018, &H2019, &H201C, &H201D, &HA9, &HAC, &H23A9, &H2321}},
    {"}", New Integer() {&H2014, &H2DC, &H203A, &HF09D, &H2044, &H221E, &H2666}},
    {"#", New Integer() {&H2194, &H23AE}},
    {"\x1621", New Integer() {&H2190, &H2033, &H2265, &HD7, &H221D}},
    {"`", New Integer() {&H2261, &H2248, &H2211}},
    {",", New Integer() {&H2026, &H23D0}},
    {".", New Integer() {&H23AF, &H21B5}},
    {"`#", New Integer() {&H2135}},
    {"`^", New Integer() {&H211C, &H2118}},
    {"_", New Integer() {&H2297}},
    {"=", New Integer() {&H2229}},
    {"", New Integer() {&H222A, &HFF}},
    {"0", New Integer() {&H2283}},
    {"1", New Integer() {&H2287}},
    {"2", New Integer() {&H2284}},
    {"3", New Integer() {&H2282}},
    {"4", New Integer() {&H2286}},
    {"5", New Integer() {&H2208}},
    {"6", New Integer() {&H2209}},
    {"7", New Integer() {&H2220}},
    {"8", New Integer() {&H2207}},
    {"H", New Integer() {&H220F, &H21D4}},
    {"x", New Integer() {&H221A, &H21D0}},
    {"Y#", New Integer() {&H22C5}},
    {"j", New Integer() {&H2228}},
    {"S", New Integer() {&H21D1, &H21D2}},
    {"D", New Integer() {&H21D3}},
    {"(", New Integer() {&H2329}},
    {")", New Integer() {&HAE}},
    {"'", New Integer() {&H2122}},
    {"t", New Integer() {&H23A2, &H222B, &H23A0}},
    {"b", New Integer() {&H2320, &H23A4}},
    {"v", New Integer() {&H239F}}
    } '{"9", New Integer() {&HAE}},
    Shared Arr3 As New Dictionary(Of String, Integer()) From {
{"l", New Integer() {&H21, &H23, &H51}},
{"d", New Integer() {&H24}},
{"*", New Integer() {&H25}},
{"h", New Integer() {&H26, &H53, &H2026, &H2030}},
{"p", New Integer() {&H27}},
{"t", New Integer() {&H29, &H56, &H6E}},
{"y", New Integer() {&H2C, &H71}},
{"Y", New Integer() {&H2E}},
{"m", New Integer() {&H2F}},
{"`", New Integer() {&H31, &H38, &H39}},
{"n", New Integer() {&H47}},
{"", New Integer() {&H52, &H5E, &H5F, &H201E, &H2020, &HF08E}},
{"llh", New Integer() {&H54}},
{"b", New Integer() {&H55}},
{"v", New Integer() {&H57}},
{"f", New Integer() {&H58, &H2019}},
{"q", New Integer() {&H59}},
{"s", New Integer() {&H5A, &H5B}},
{"$", New Integer() {&H5C}},
{"{lr~aHiymi", New Integer() {&H69}},
{"{ll~`hi {lr~aHoma`ni", New Integer() {&H6A}},
{"bisomi", New Integer() {&H6B}},
{"k", New Integer() {&H6C, &H201A, &HF081}},
{"E", New Integer() {&H192, &H201D}}
}
    Shared Arr4 As New Dictionary(Of String, Integer()) From {
{" ", New Integer() {&H20}},
{"\x1751", New Integer() {&H22, &H33}},
{"\x1754", New Integer() {&H23, &H34}},
{"^", New Integer() {&H24, &HFA, &HFB, &HFC, &HFD, &HFE, &HFF}},
{"K", New Integer() {&H25, &H26, &H35, &H37, &H38, &H39, &H3A, &H3B, &H3C, &H3E, &H3F, &H40, &H41, &H42, &H43}},
{"\x1750", New Integer() {&H28}},
{"R", New Integer() {&H29}},
{"B", New Integer() {&H2A}},
{"F[", New Integer() {&H47, &H49, &H4A, &H4B, &H4C, &H4D, &H4E, &H4F, &H50, &H52}},
{"F", New Integer() {&H54, &H55, &H56, &H57, &H58, &H59, &H5A, &H5B, &H5C, &H5D, &HB2, &HB3, &HB4, &HB5, &HB6, &HB7, &HB8, &HB9, &HBA, &HBB}},
{"~", New Integer() {&H5F, &H61, &H62, &H63, &H64, &H65, &H67, &H68, &H69, &H6A, &H6B, &H6D, &H6E, &H6F, &H70}},
{"~F", New Integer() {&H72, &H74, &H75, &H76, &H77, &H78, &H79, &H7B, &H7C, &H7D, &H7E, &H201A, &H192, &H2026, &H2020, &H2021, &H2C6, &H2030}},
{"~u", New Integer() {&HF08E, &HF08F, &HF090, &H2018, &H2019, &H201C, &H201D, &H2022, &H2013, &H2014, &H2122, &H161, &H203A, &H153, &HF09D}},
{"~a", New Integer() {&HF09E, &HA0, &HA1, &HA2, &HA3, &HA4, &HA5, &HA6, &HA7, &HA8, &HA9, &HAA, &HAC, &H2212, &HAE, &HAF, &HB0}},
{"i", New Integer() {&HBD, &HBE, &HBF, &HC2, &HC3, &HC4, &HC5, &HC6, &HC7, &HC8, &HC9, &HCA, &HCB, &HCC, &HCD, &HCE, &HCF}},
{"N", New Integer() {&HD0, &HD1, &HD2, &HD3, &HD4, &HD6, &HD7, &HD8, &HEA, &HEB, &HEC, &HED, &HEE}},
{"u", New Integer() {&HD9, &HDA, &HDB, &HDC, &HDD, &HDE, &HDF, &HE0, &HE1, &HE2, &HE3, &HE4, &HE5, &HE6, &HE7, &HE8, &HE9}},
{"o", New Integer() {&HEF, &HF1, &HF2, &HF3, &HF4, &HF5, &HF6, &HF7, &HF8, &HF9}}
}
    'combine with alef
    Shared Arr5 As New Dictionary(Of String, Integer()) From {
    {" ", New Integer() {&H20}},
    {"^", New Integer() {&H21, &HF0AF}},
    {"`^", New Integer() {&H23}},
    {"{", New Integer() {&H24}},
    {"o", New Integer() {&H26, &H27, &H28}},
    {"[", New Integer() {&H2D, &H2E, &H2F}},
    {"a+", New Integer() {&H30}},
    {"-", New Integer() {&H31}},
    {"`", New Integer() {&H32, &H33, &H34}},
    {"u[", New Integer() {&H36, &H37, &H38, &H39}},
    {"~u^", New Integer() {&H3A, &H3B, &H3C}},
    {"~N", New Integer() {&H3F, &H40, &H41}},
    {"~u[", New Integer() {&H42}},
    {"~a[", New Integer() {&H43}},
    {"oa", New Integer() {&H46}},
    {"~a^", New Integer() {&H48}},
    {"a^", New Integer() {&H49}},
    {"ia", New Integer() {&H4B, &H4C}},
    {"oi", New Integer() {&H4D}},
    {"""", New Integer() {&H4F}},
    {"iu", New Integer() {&H54}},
    {"aa", New Integer() {&H55, &H56}},
    {"~ia", New Integer() {&H58}},
    {"ai", New Integer() {&H5A}},
    {"~iu", New Integer() {&H5B}},
    {"ii", New Integer() {&H5C}},
    {"a{", New Integer() {&H5D}},
    {"~a{", New Integer() {&H5E}},
    {"au", New Integer() {&H5F}},
    {"~aa", New Integer() {&H60, &HA9}},
    {"~ii", New Integer() {&H62}},
    {"~au", New Integer() {&H63, &HAA}},
    {"aN", New Integer() {&H64}},
    {"i{", New Integer() {&H65}},
    {"ao", New Integer() {&H68}},
    {"a", New Integer() {&H6D, &H6E, &H6F, &H70, &H71, &H72, &H73, &H74, &H75, &H76, &H77, &H78, &H79, &H7A, &H7B, &H7C, &H7D, &H7E, &HF080, &HF08D, &H161, &H178}},
    {"~a", New Integer() {&H153, &HF09E}},
    {"\x1753", New Integer() {&HA0}},
    {"\x1755", New Integer() {&HA1}},
    {"\x1752", New Integer() {&HA2}},
    {"i]", New Integer() {&HA3, &HA4, &HA5}},
    {"~ai", New Integer() {&HAB}},
    {"i~ai", New Integer() {&HAC}},
    {"~i~ai", New Integer() {&HB0}}
    }
    Shared Arr7 As New Dictionary(Of String, Integer()) From {
    {"-", New Integer() {&H31}},
    {"a", New Integer() {&H76, &H79, &H7D, &HF08D}}
    }
    Class CompareChar
        Implements IComparer(Of ArrayList)
        Public Function Compare(x As ArrayList, y As ArrayList) As Integer Implements IComparer(Of ArrayList).Compare
            'If x(3) = y(3) Then Return 0
            'Return If(x(3) > y(3), 1, -1)
            If CSng(x(1)) = CSng(y(1)) Then
                If CSng(x(0)) = CSng(y(0)) Then
                    Return 1
                End If
                Return If(CSng(x(0)) > CSng(y(0)), -1, 1)
            End If
            Return If(CSng(x(1)) > CSng(y(1)), -1, 1)
        End Function
    End Class
    Class TextExtractionStrategy
        Implements iTextSharp.text.pdf.parser.ITextExtractionStrategy
        Class Chunk
            Public Str As String
            Public FontName As String
            Public Start As iTextSharp.text.pdf.parser.Vector
            Public Finish As iTextSharp.text.pdf.parser.Vector
            Public Sub New(NewStr As String, NewFontName As String, NewStart As iTextSharp.text.pdf.parser.Vector, NewFinish As iTextSharp.text.pdf.parser.Vector)
                Str = NewStr
                FontName = NewFontName
                Start = NewStart
                Finish = NewFinish
            End Sub
        End Class
        Dim Chunks As New Generic.List(Of Chunk)

        Public Sub BeginTextBlock() Implements iTextSharp.text.pdf.parser.IRenderListener.BeginTextBlock

        End Sub

        Public Sub EndTextBlock() Implements iTextSharp.text.pdf.parser.IRenderListener.EndTextBlock

        End Sub

        Public Sub RenderImage(renderInfo As iTextSharp.text.pdf.parser.ImageRenderInfo) Implements iTextSharp.text.pdf.parser.IRenderListener.RenderImage

        End Sub

        Public Sub RenderText(renderInfo As iTextSharp.text.pdf.parser.TextRenderInfo) Implements iTextSharp.text.pdf.parser.IRenderListener.RenderText
            For Count As Integer = 0 To renderInfo.GetCharacterRenderInfos.Count - 1
                Dim Segment As iTextSharp.text.pdf.parser.LineSegment = renderInfo.GetCharacterRenderInfos(Count).GetBaseline()
                Chunks.Add(New Chunk(renderInfo.GetCharacterRenderInfos(Count).GetText(), renderInfo.GetFont().PostscriptFontName, Segment.GetStartPoint(), Segment.GetEndPoint()))
            Next
        End Sub

        Public Function GetResultantText() As String Implements iTextSharp.text.pdf.parser.ITextExtractionStrategy.GetResultantText
            Dim Str As String = String.Empty
            Chunks.Sort(Function(First As Chunk, Second As Chunk)
                            If First.Finish.Item(iTextSharp.text.pdf.parser.Vector.I2) = Second.Finish.Item(iTextSharp.text.pdf.parser.Vector.I2) Then
                                Return If(First.Finish.Item(iTextSharp.text.pdf.parser.Vector.I1) > Second.Finish.Item(iTextSharp.text.pdf.parser.Vector.I1), -1, 1)
                            End If
                            Return If(First.Finish.Item(iTextSharp.text.pdf.parser.Vector.I2) > Second.Finish.Item(iTextSharp.text.pdf.parser.Vector.I2), -1, 1)

                            Dim Vec1 As iTextSharp.text.pdf.parser.Vector = First.Finish.Subtract(First.Start)
                            If Vec1.Length = 0 Then
                                Vec1 = New iTextSharp.text.pdf.parser.Vector(1, 0, 0)
                            Else
                                Vec1 = Vec1.Normalize()
                            End If
                            Dim Vec2 As iTextSharp.text.pdf.parser.Vector = Second.Finish.Subtract(Second.Start)
                            If Vec2.Length = 0 Then
                                Vec2 = New iTextSharp.text.pdf.parser.Vector(1, 0, 0)
                            Else
                                Vec2 = Vec1.Normalize()
                            End If
                            Dim Mag1 As Integer = CInt(Math.Atan2(Vec1(iTextSharp.text.pdf.parser.Vector.I2), Vec1(iTextSharp.text.pdf.parser.Vector.I1)) * 1000)
                            Dim Mag2 As Integer = CInt(Math.Atan2(Vec2(iTextSharp.text.pdf.parser.Vector.I2), Vec2(iTextSharp.text.pdf.parser.Vector.I1)) * 1000)
                            If Mag1 = Mag2 Then
                                Dim DistPerp1 As Integer = CInt(First.Start.Subtract(New iTextSharp.text.pdf.parser.Vector(0, 0, 1)).Cross(Vec1).Item(iTextSharp.text.pdf.parser.Vector.I3))
                                Dim DistPerp2 As Integer = CInt(Second.Start.Subtract(New iTextSharp.text.pdf.parser.Vector(0, 0, 1)).Cross(Vec2).Item(iTextSharp.text.pdf.parser.Vector.I3))
                                If DistPerp1 = DistPerp2 Then
                                    Return If(Vec1.Dot(First.Start) < Vec2.Dot(Second.Start), -1, 1)
                                Else
                                    Return If(DistPerp1 < DistPerp2, -1, 1)
                                End If
                            Else
                                Return If(Mag1 < Mag2, -1, 1)
                            End If
                        End Function)
            For Count As Integer = 0 To Chunks.Count - 1
                Dim CurDict As Dictionary(Of String, Integer()) = Nothing
                If Chunks(Count).FontName = "BAMCIC+HQPB1" Then
                    CurDict = Arr1
                ElseIf Chunks(Count).FontName = "BAMCGB+HQPB2" Then
                    CurDict = Arr2
                ElseIf Chunks(Count).FontName = "BAMCJD+HQPB3" Then
                    CurDict = Arr3
                ElseIf Chunks(Count).FontName = "BAMCGA+HQPB4" Then
                    CurDict = Arr4
                ElseIf Chunks(Count).FontName = "BAMCHB+HQPB5" Then
                    CurDict = Arr5
                ElseIf Chunks(Count).FontName = "BAMDFB+HQPB7" Then
                    CurDict = Arr7
                End If
                If Not CurDict Is Nothing Then
                    For Each KeyValue As Collections.Generic.KeyValuePair(Of String, Integer()) In CurDict
                        If Array.IndexOf(KeyValue.Value, System.Text.Encoding.Unicode.GetBytes(Chunks(Count).Str)(0) + 256 * System.Text.Encoding.Unicode.GetBytes(Chunks(Count).Str)(1)) <> -1 Then
                            Str += KeyValue.Key
                        End If
                    Next
                End If
            Next
            Return Str
        End Function
    End Class
    Shared Function ParseQuran() As String
        Dim Reader As New iTextSharp.text.pdf.PdfReader("..\..\..\IslamMetadata\warsh.pdf")
        For Cnt As Integer = 0 To Reader.NumberOfPages - 1
            iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(Reader, Cnt + 1, New TextExtractionStrategy)

        Next
        'Tf - Font, Size
        '"/Resources/Font/TT1/BaseFont"
        'Tj Or "TJ"
        Dim Lines As Byte() = IO.File.ReadAllBytes("..\..\..\IslamMetadata\warsh.csv")
        Dim LinesOrder As String() = IO.File.ReadAllLines("..\..\..\IslamMetadata\warshp.txt", System.Text.Encoding.UTF8)
        Dim CharTable As String = String.Empty
        For Count As Integer = 0 To 1000 'LinesOrder.Length - 1
            CharTable += (New System.Text.UnicodeEncoding).GetString(System.Text.Encoding.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.Unicode, (New System.Text.UTF8Encoding).GetBytes((StrReverse(LinesOrder(Count)) + " ").ToCharArray())))
        Next
        Dim Bytes As Char() = CharTable.ToCharArray()
        Dim Chars As New ArrayList
        Dim Pos As Integer = 0
        Dim bNextPage As Boolean = False
        For Count As Integer = 0 To 360 'Lines.Length - 1
            Dim Vals(11) As String
            Vals(0) = System.Text.Encoding.Default.GetString(Lines, Pos, Array.IndexOf(Lines, CByte(Asc(","c)), Pos) - Pos).Trim(""""c)
            Pos = Array.IndexOf(Lines, CByte(Asc(","c)), Pos) + 1
            Vals(1) = System.Text.Encoding.Default.GetString(Lines, Pos, Array.IndexOf(Lines, CByte(Asc(","c)), Pos) - Pos)
            Pos = Array.IndexOf(Lines, CByte(Asc(","c)), Pos) + 1
            Vals(2) = System.Text.Encoding.Default.GetString(Lines, Pos, Array.IndexOf(Lines, CByte(Asc(","c)), Pos) - Pos)
            Pos = Array.IndexOf(Lines, CByte(Asc(","c)), Pos) + 1
            Vals(3) = System.Text.Encoding.Default.GetString(Lines, Pos, Array.IndexOf(Lines, CByte(Asc(","c)), Pos) - Pos)
            Pos = Array.IndexOf(Lines, CByte(Asc(","c)), Pos) + 1
            Vals(4) = System.Text.Encoding.Default.GetString(Lines, Pos, Array.IndexOf(Lines, CByte(Asc(","c)), Pos) - Pos)
            Pos = Array.IndexOf(Lines, CByte(Asc(","c)), Pos) + 1
            Vals(5) = System.Text.Encoding.Default.GetString(Lines, Pos, Array.IndexOf(Lines, CByte(Asc(","c)), Pos) - Pos)
            Pos = Array.IndexOf(Lines, CByte(Asc(","c)), Pos) + 1
            Vals(6) = System.Text.Encoding.Default.GetString(Lines, Pos, Array.IndexOf(Lines, CByte(Asc(","c)), Pos) - Pos)
            Pos = Array.IndexOf(Lines, CByte(Asc(","c)), Pos) + 1
            Vals(7) = System.Text.Encoding.Default.GetString(Lines, Pos, Array.IndexOf(Lines, CByte(Asc(","c)), Pos) - Pos)
            Pos = Array.IndexOf(Lines, CByte(Asc(","c)), Pos) + 1
            Vals(8) = System.Text.Encoding.Default.GetString(Lines, Pos, Array.IndexOf(Lines, CByte(Asc(","c)), Pos) - Pos)
            Pos = Array.IndexOf(Lines, CByte(Asc(","c)), Pos) + 1
            Vals(9) = System.Text.Encoding.Default.GetString(Lines, Pos, Array.IndexOf(Lines, CByte(Asc(","c)), Pos) - Pos)
            Pos = Array.IndexOf(Lines, CByte(Asc(","c)), Pos) + 1
            Vals(10) = System.Text.Encoding.Default.GetString(Lines, Pos, Array.IndexOf(Lines, CByte(Asc(","c)), Pos) - Pos)
            Pos = Array.IndexOf(Lines, CByte(Asc(","c)), Pos) + 1
            If Vals(0) = "BAMCBO+TimesNewRomanPSMT" Then
                Vals(11) = System.Text.Encoding.Unicode.GetString(System.Text.Encoding.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.Unicode, Lines, Pos, Array.IndexOf(Lines, CByte(Asc(vbCr)), Pos) - Pos)).Trim(""""c)
                Pos = Array.IndexOf(Lines, CByte(Asc(vbLf)), Pos) + 1
                Dim Idx As Integer = 0
                If Vals(11).Replace(" ", String.Empty) <> String.Empty Or Vals(11) = " " Then
                    For SubCount = Vals(11).Length - 1 To 0 Step -1
                        Idx = Array.IndexOf(Bytes, Vals(11).Chars(SubCount))
                        Bytes(Idx) = ChrW(0)
                    Next
                End If
                If Not bNextPage And Vals(11).Length > 1 And Vals(11).Length < 10 And Vals(11).Replace(" ", String.Empty) = String.Empty Then
                    'Idx = Array.IndexOf(Bytes, " "c)
                    'Bytes(Idx) = ChrW(0)
                    If Chars.Count <> 0 Then Chars.Add(New ArrayList From {CSng(Vals(7)), CSng(Vals(8)) - If(CSng(Vals(8)) - 0.6571 = CSng(CType(Chars(Chars.Count - 1), ArrayList)(1)), 0.6571, 0.657), Vals(11), Idx})
                End If
                If Vals(11) = " " Then
                    If Chars.Count <> 0 Then Chars.Add(New ArrayList From {CSng(Vals(7)), CSng(Vals(8)) - If(CSng(Vals(8)) - 0.6571 = CSng(CType(Chars(Chars.Count - 1), ArrayList)(1)), 0.6571, 0.657), Vals(11), Idx})
                End If
            Else
                Dim CurDict As Dictionary(Of String, Integer()) = Nothing
                If Vals(0) = "BAMCIC+HQPB1" Then
                    CurDict = Arr1
                ElseIf Vals(0) = "BAMCGB+HQPB2" Then
                    CurDict = Arr2
                ElseIf Vals(0) = "BAMCJD+HQPB3" Then
                    CurDict = Arr3
                ElseIf Vals(0) = "BAMCGA+HQPB4" Then
                    CurDict = Arr4
                ElseIf Vals(0) = "BAMCHB+HQPB5" Then
                    CurDict = Arr5
                ElseIf Vals(0) = "BAMDFB+HQPB7" Then
                    CurDict = Arr7
                Else
                    Vals(11) = System.Text.Encoding.Unicode.GetString(System.Text.Encoding.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.Unicode, Lines, Pos, Array.IndexOf(Lines, CByte(Asc(vbCr)), Pos) - Pos)).Trim(""""c)
                    Pos = Array.IndexOf(Lines, CByte(Asc(vbLf)), Pos) + 1
                    If Vals(11).Replace(" ", String.Empty) <> String.Empty Or Vals(11) = " " Then
                        For SubCount = Vals(11).Length - 1 To 0 Step -1
                            'Dim Idx As Integer = Array.IndexOf(Bytes, Vals(11).Chars(SubCount))
                            'If Idx <> -1 Then Bytes(Idx) = ChrW(0)
                        Next
                    End If
                End If
                If Vals(0) = "TraditionalArabic" Then
                    bNextPage = True
                End If
                If Not CurDict Is Nothing Then
                    If bNextPage Then
                        'Bytes(Array.IndexOf(Bytes, " "c)) = ChrW(0)
                        'MsgBox(StrReverse(Vals(11)) + "/" + CStr(Array.IndexOf(Bytes, " "c)) + "/" + CStr(Array.LastIndexOf(Bytes, ChrW(0), Array.IndexOf(Bytes, " "c))) + "/" + String.Join(String.Empty, Array.ConvertAll(Bytes, Function(C As Char) CStr(C)), Array.LastIndexOf(Bytes, ChrW(0), Array.IndexOf(Bytes, " "c)) + 1, Array.IndexOf(Bytes, " "c) - Array.LastIndexOf(Bytes, " "c, Array.IndexOf(Bytes, " "c))))
                        bNextPage = False
                        If Bytes(Array.IndexOf(Bytes, " "c) - 1) = ChrW(&H2329) Or Bytes(Array.IndexOf(Bytes, " "c) - 1) = ChrW(&HAE) Then

                        End If
                    End If
                    Vals(11) = System.Text.Encoding.Unicode.GetString(System.Text.Encoding.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.Unicode, Lines, Pos, Array.IndexOf(Lines, CByte(Asc(vbCr)), Pos) - Pos)).Trim(""""c)
                    Pos = Array.IndexOf(Lines, CByte(Asc(vbLf)), Pos) + 1
                    Dim privateFonts As New System.Drawing.Text.PrivateFontCollection()
                    privateFonts.AddFontFile("..\..\..\IslamMetadata\" + Vals(0) + ".ttf")
                    Dim f As New Font(privateFonts.Families(0), CSng(Vals(2)), GraphicsUnit.World)
                    Dim i As New System.Drawing.Bitmap(595, 842)
                    i.SetResolution(72, 72)
                    '8.26 width = 595, 11.69 height = 842, 72 DPI
                    Dim g As Graphics = Graphics.FromImage(i)
                    g.PageUnit = GraphicsUnit.Point
                    g.PageScale = 72 / 100
                    Dim Fmat As New StringFormat
                    Fmat.FormatFlags = StringFormatFlags.MeasureTrailingSpaces ' StringFormatFlags.DirectionRightToLeft ' Or StringFormatFlags.MeasureTrailingSpaces
                    Dim Ranges(Vals(11).Length - 1) As CharacterRange
                    For SubCount As Integer = Vals(11).Length - 1 To 0 Step -1
                        Ranges(SubCount) = New CharacterRange(0, SubCount + 1)
                    Next
                    Fmat.SetMeasurableCharacterRanges(Ranges)
                    Dim Rgs As Region() = g.MeasureCharacterRanges((Vals(11)), f, New RectangleF(0, 0, 595, 842), Fmat)
                    For SubCount = Vals(11).Length - 1 To 0 Step -1
                        If Vals(11).Chars(SubCount) <> " " Then
                            For Each KeyValue As Collections.Generic.KeyValuePair(Of String, Integer()) In CurDict
                                If Array.IndexOf(KeyValue.Value, System.Text.Encoding.Unicode.GetBytes(Vals(11).Chars(SubCount))(0) + 256 * System.Text.Encoding.Unicode.GetBytes(Vals(11).Chars(SubCount))(1)) <> -1 Then
                                    ' If SubCount = 0 And (KeyValue.Key.Chars(0) = "a" Or KeyValue.Key.Chars(0) = "i" Or KeyValue.Key.Chars(0) = "i" Or KeyValue.Key.Chars(0) = "u" Or KeyValue.Key.Chars(0) = "o" Or KeyValue.Key.Chars(0) = "F" Or KeyValue.Key.Chars(0) = "N" Or KeyValue.Key.Chars(0) = "K" Or KeyValue.Key.Chars(0) = "~" Or KeyValue.Key.Chars(0) = "`") Then

                                    'End If
                                    Dim Idx As Integer = Array.LastIndexOf(Bytes, Vals(11).Chars(SubCount), Array.IndexOf(Bytes, " "c))
                                    'If Idx = -1 Then Idx = Array.IndexOf(Bytes, Vals(11).Chars(SubCount))
                                    Chars.Add(New ArrayList From {If(SubCount = Vals(11).Length - 1, CSng(Vals(7)), CSng(Vals(9)) + Rgs(SubCount).GetBounds(g).Width * 72 / 100), CSng(Vals(8)), KeyValue.Key, Idx})
                                    'Bytes(Idx) = ChrW(0)
                                    '(CSng(Vals(7)) - CSng(Vals(9))) / Rgs(Vals(11).Length - 1).GetBounds(g).Width *
                                    Exit For
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        Next
        Dim SortChars As ArrayList() = CType(Chars.ToArray(GetType(ArrayList)), ArrayList())
        Array.Sort(SortChars, New CompareChar)
        ParseQuran = Nothing
        For Count As Integer = 0 To SortChars.Length - 1
            If CStr(SortChars(Count)(2)) = "A" AndAlso CStr(SortChars(Count + 1)(2)) = "{" Then
            ElseIf CStr(SortChars(Count)(2)) = "llh" Then
            ElseIf CStr(SortChars(Count)(2)) = "i~ai" Or CStr(SortChars(Count)(2)) = "~i~ai" Then
                If CStr(SortChars(Count - 1)(2)) = "llh" Then
                    ParseQuran += CStr(SortChars(Count - 1)(2)).Chars(0) + CStr(SortChars(Count)(2)).Chars(0)
                    If CStr(SortChars(Count)(2)).Chars(0) = "~" Then
                        ParseQuran += CStr(SortChars(Count)(2)).Chars(1)
                    End If
                    ParseQuran += CStr(SortChars(Count - 1)(2)).Chars(1)
                    ParseQuran += CStr(SortChars(Count)(2)).Chars(If(CStr(SortChars(Count)(2)).Chars(0) = "~", 2, 1)) + CStr(SortChars(Count)(2)).Chars(If(CStr(SortChars(Count)(2)).Chars(0) = "~", 3, 2))
                    ParseQuran += CStr(SortChars(Count - 1)(2)).Chars(2)
                    ParseQuran += CStr(SortChars(Count)(2)).Chars(If(CStr(SortChars(Count)(2)).Chars(0) = "~", 4, 3))
                Else
                    ParseQuran += CStr(SortChars(Count)(2)).Chars(0)
                    If CStr(SortChars(Count)(2)).Chars(0) = "~" Then
                        ParseQuran += CStr(SortChars(Count)(2)).Chars(1)
                    End If
                    ParseQuran += CStr(SortChars(Count + 1)(2))
                    ParseQuran += CStr(SortChars(Count)(2)).Chars(If(CStr(SortChars(Count)(2)).Chars(0) = "~", 2, 1)) + CStr(SortChars(Count)(2)).Chars(If(CStr(SortChars(Count)(2)).Chars(0) = "~", 3, 2))
                    ParseQuran += CStr(SortChars(Count + 2)(2))
                    ParseQuran += CStr(SortChars(Count)(2)).Chars(If(CStr(SortChars(Count)(2)).Chars(0) = "~", 4, 3))
                    Count += 2
                End If
            ElseIf CStr(SortChars(Count)(2)) <> String.Empty AndAlso CStr(SortChars(Count)(2)).Length >= If(CStr(SortChars(Count)(2)).Chars(0) = "~", 3, 2) AndAlso ("aiuo".IndexOf(CStr(SortChars(Count)(2)).Chars(1)) <> -1 And "aiuo".IndexOf(CStr(SortChars(Count)(2)).Chars(If(CStr(SortChars(Count)(2)).Chars(0) = "~", 2, 0))) <> -1) Then
                '"oa", "ia", "oi", "iu", "aa", "~ia", "ai", "~iu", "ii", "au", "~aa", "~ii", "~au", "ao", "~ai"
                If CStr(SortChars(Count - 1)(2)) = "llh" Then
                    ParseQuran += CStr(SortChars(Count - 1)(2)).Chars(0) + CStr(SortChars(Count - 1)(2)).Chars(1) + CStr(SortChars(Count)(2)).Chars(0)
                    If CStr(SortChars(Count)(2)).Chars(0) = "~" Then
                        ParseQuran += CStr(SortChars(Count)(2)).Chars(1)
                    End If
                    ParseQuran += CStr(SortChars(Count - 1)(2)).Chars(2)
                    ParseQuran += CStr(SortChars(Count)(2)).Chars(If(CStr(SortChars(Count)(2)).Chars(0) = "~", 2, 1))
                Else
                    ParseQuran += CStr(SortChars(Count)(2)).Chars(0)
                    If CStr(SortChars(Count)(2)).Chars(0) = "~" Then
                        ParseQuran += CStr(SortChars(Count)(2)).Chars(1)
                    End If
                    ParseQuran += CStr(SortChars(Count + 1)(2))
                    ParseQuran += CStr(SortChars(Count)(2)).Chars(If(CStr(SortChars(Count)(2)).Chars(0) = "~", 2, 1))
                    Count += 1
                End If
            Else
                ParseQuran += CStr(SortChars(Count)(2))
            End If
        Next
        MsgBox(ParseQuran)
    End Function
End Class
