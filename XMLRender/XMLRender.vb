Option Explicit On
Option Strict On
<CLSCompliant(True)>
Public Interface PortableSettings
    ReadOnly Property CacheDirectory As String
    ReadOnly Property Resources As KeyValuePair(Of String, String())()
    ReadOnly Property FuncLibs As String()
    Function GetTemplatePath(Selector As String) As String
    Function GetFilePath(ByVal Path As String) As String
    Function GetUName(Character As Char) As String
    Function GetResourceString(baseName As String, resourceKey As String) As String
End Interface
<CLSCompliant(True)>
Public Interface PortableFileIO
    Function LoadStream(FilePath As String) As Task(Of IO.Stream)
    Function SaveStream(FilePath As String, Stream As IO.Stream) As Task
    Function GetDirectoryFiles(Path As String) As Task(Of String())
    Function CombinePath(ParamArray Paths() As String) As String
    Function DeleteFile(FilePath As String) As Task
    Function PathExists(Path As String) As Task(Of Boolean)
    Function CreateDirectory(Path As String) As Task
    Function PathGetLastWriteTimeUtc(Path As String) As Task(Of Date)
    Function PathSetLastWriteTimeUtc(Path As String, Time As Date) As Task
End Interface
<CLSCompliant(True)>
Public Class PortableMethods
    Public FileIO As PortableFileIO
    Public Settings As PortableSettings
    Public DiskCache As DiskCacheUtility
    Public Sub New(NewFileIO As PortableFileIO, NewSettings As PortableSettings)
        FileIO = NewFileIO
        Settings = NewSettings
        DiskCache = New DiskCacheUtility
    End Sub
    Public Async Function Init() As Task
        Await DiskCache.Init(Me)
    End Function
    Public Class DiskCacheUtility
        Private _PortableMethods As PortableMethods
        Public Async Function Init(NewPortableMethods As PortableMethods) As Task
            _PortableMethods = NewPortableMethods
            Await GetCacheDirectory()
        End Function
        Async Function GetCacheDirectory() As Task(Of String)
            Dim Path As String
            Path = _PortableMethods.FileIO.CombinePath(_PortableMethods.Settings.CacheDirectory, "DiskCache")
            If Not (Await _PortableMethods.FileIO.PathExists(Path)) Then Await _PortableMethods.FileIO.CreateDirectory(Path)
            Return Path
        End Function
        Public Async Function GetCacheItem(ByVal Name As String, ByVal ModifiedUtc As Date) As Task(Of Byte())
            If Not Await _PortableMethods.FileIO.PathExists(_PortableMethods.FileIO.CombinePath(Await GetCacheDirectory(), Name)) OrElse
            ModifiedUtc > Await _PortableMethods.FileIO.PathGetLastWriteTimeUtc(_PortableMethods.FileIO.CombinePath(Await GetCacheDirectory(), Name)) Then Return Nothing
            Dim File As IO.Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.FileIO.CombinePath(Await GetCacheDirectory(), Name))
            Dim Bytes(CInt(File.Length) - 1) As Byte
            File.Read(Bytes, 0, CInt(File.Length))
            File.Dispose()
            Return Bytes
        End Function
        Public Async Function GetCacheItemPath(ByVal Name As String, ByVal ModifiedUtc As Date) As Task(Of String)
            If Not Await _PortableMethods.FileIO.PathExists(_PortableMethods.FileIO.CombinePath(Await GetCacheDirectory(), Name)) OrElse
            ModifiedUtc > Await _PortableMethods.FileIO.PathGetLastWriteTimeUtc(_PortableMethods.FileIO.CombinePath(Await GetCacheDirectory(), Name)) Then Return String.Empty
            Return _PortableMethods.FileIO.CombinePath(Await GetCacheDirectory(), Name)
        End Function
        Public Async Function CacheItem(ByVal Name As String, ByVal LastModifiedUtc As Date, ByVal Data() As Byte) As Task
            Dim File As New IO.MemoryStream
            File.Write(Data, 0, Data.Length)
            Await _PortableMethods.FileIO.SaveStream(_PortableMethods.FileIO.CombinePath(Await GetCacheDirectory(), Name), File)
            File.Dispose()
            If LastModifiedUtc = DateTime.MinValue Then LastModifiedUtc = DateTime.Now
            Await _PortableMethods.FileIO.PathSetLastWriteTimeUtc(_PortableMethods.FileIO.CombinePath(Await GetCacheDirectory(), Name), LastModifiedUtc)
        End Function
        Public Async Function GetCacheItems() As Task(Of String())
            Return Await _PortableMethods.FileIO.GetDirectoryFiles(Await GetCacheDirectory())
        End Function
        Public Async Function DeleteUnusedCacheItems(ByVal ActiveNames() As String) As Task
            Dim Files() As String = Await _PortableMethods.FileIO.GetDirectoryFiles(Await GetCacheDirectory())
            Dim Count As Integer
            For Count = 0 To Files.Length - 1
                If Array.IndexOf(ActiveNames, Files(Count)) = -1 Then Await DeleteCacheItem(Files(Count))
            Next
        End Function
        Public Async Function DeleteCacheItem(ByVal Name As String) As Task
            Await _PortableMethods.FileIO.DeleteFile(Name)
        End Function
    End Class
    Public Function LoadResourceString(resourceKey As String) As String
        LoadResourceString = Nothing
        If resourceKey Is Nothing Then Return Nothing
        For Each Pair In Settings.Resources
            If Linq.Enumerable.TakeWhile(Pair.Value, Function(Str As String) Not (Str = resourceKey Or resourceKey.StartsWith(Str + "_"))).Count() <> Pair.Value.Count() Then
                LoadResourceString = Settings.GetResourceString(Pair.Key, resourceKey)
            End If
        Next
        If LoadResourceString = Nothing And Not resourceKey.EndsWith("_") Then
            LoadResourceString = String.Empty
            System.Diagnostics.Debug.WriteLine("  <data name=""" + resourceKey + """ xml:space=""preserve"">" + vbCrLf + "    <value>" + System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(resourceKey, ".*_", String.Empty), "(.+?)([A-Z])", "$1 $2") + "</value>" + vbCrLf + "  </data>")
        End If
    End Function
    Public Async Function WordFileToResource(WordFilePath As String, ResFilePath As String) As Task
        Dim W4WLines As String() = Await ReadAllLines(WordFilePath)
        Dim Stream As IO.Stream = Await FileIO.LoadStream(ResFilePath)
        Dim XMLDoc As Xml.Linq.XDocument = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        For Count = 0 To W4WLines.Length - 1
            Dim Words As String() = W4WLines(Count).Substring(W4WLines(Count).IndexOf("="c) + 1).Split("|"c)
            For SubCount = 0 To Words.Length - 1
                Dim NewData As New Xml.Linq.XElement("data")
                NewData.Add(New Xml.Linq.XAttribute("name", W4WLines(Count).Substring(0, W4WLines(Count).IndexOf("="c)) + "." + CStr(SubCount + 1)))
                NewData.Add(New Xml.Linq.XAttribute(Xml.Linq.XNamespace.Xml + "space", "preserve"))
                Dim Inner As New Xml.Linq.XElement("value")
                Inner.Value = Words(SubCount)
                NewData.Add(Inner)
                XMLDoc.Root.Add(NewData)
            Next
        Next
        Dim MemStream As New IO.MemoryStream
        XMLDoc.Save(MemStream)
        Await FileIO.SaveStream(ResFilePath, MemStream)
        MemStream.Dispose()
    End Function
    Public Async Function ResourceToWordFile(ResFilePath As String, WordFilePath As String) As Task
        Dim W4WLines As New Dictionary(Of String, List(Of String))
        Dim Stream As IO.Stream = Await FileIO.LoadStream(ResFilePath)
        Dim XMLDoc As Xml.Linq.XDocument = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim AllNodes As Xml.Linq.XElement() = Linq.Enumerable.Where(XMLDoc.Root.Elements, Function(elem) elem.Name = "data" And Not elem.Attribute("name") Is Nothing).ToArray()
        For Each Item As Xml.Linq.XElement In AllNodes
            If System.Text.RegularExpressions.Regex.Match(Item.Attribute("name").Value, "^.*\.\d+$").Success Then
                Dim LineKey As String = Item.Attribute("name").Value.Substring(0, Item.Attribute("name").Value.IndexOf("."c))
                Dim Word As Integer = Integer.Parse(Item.Attribute("name").Value.Substring(Item.Attribute("name").Value.IndexOf("."c) + 1))
                If Not W4WLines.ContainsKey(LineKey) Then
                    W4WLines.Add(LineKey, New List(Of String))
                End If
                While W4WLines(LineKey).Count <> Word - 1
                    W4WLines(LineKey).Insert(W4WLines(LineKey).Count, String.Empty)
                End While
                W4WLines(LineKey).Insert(Word - 1, New List(Of Xml.Linq.XElement)(Item.Elements).Item(0).Value)
            End If
        Next
        Await WriteAllLines(WordFilePath, New List(Of String)(W4WLines.Select(Function(Input As KeyValuePair(Of String, List(Of String))) Input.Key + "=" + String.Join("|"c, Input.Value.ToArray()))).ToArray())
    End Function
    Public Async Function WriteAllLines(FilePath As String, Lines() As String) As Task
        Dim MemStream As New IO.MemoryStream
        Dim Writer As New IO.StreamWriter(MemStream)
        For Count As Integer = 0 To Lines.Length - 1
            Writer.WriteLine(Lines(Count))
        Next
        Writer.Flush()
        MemStream.Seek(0, SeekOrigin.Begin)
        Await FileIO.SaveStream(FilePath, MemStream)
        MemStream.Dispose()
    End Function
    Public Async Function ReadAllLines(FilePath As String) As Task(Of String())
        Dim Stream As IO.Stream = Await FileIO.LoadStream(FilePath)
        Dim Lines As String() = Utility.ReadAllLines(Stream)
        Stream.Dispose()
        Return Lines
    End Function
    Public Function GetLanguageName() As String
        Return Threading.Thread.CurrentThread.CurrentUICulture.Name
    End Function
    Public Function GetLanguageAltName() As String
        Return Threading.Thread.CurrentThread.CurrentUICulture.Parent.Name
    End Function
    Public Function GetTwoLetterISOLanguageName() As String
        Return Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName
    End Function
    Public Function GetTwoLetterISOLanguageAltName() As String
        Return Threading.Thread.CurrentThread.CurrentUICulture.Parent.TwoLetterISOLanguageName
    End Function
End Class
Public Class Utility
    Public Interface IInitClass
        Function Init(bWeb As Boolean) As Threading.Tasks.Task
        Function LookupObject(ClassName As String) As Object
        Function GetDependency() As Nullable(Of KeyValuePair(Of String, Utility.IInitClass))
    End Interface
    Public Shared Function ColorR(Clr As Integer) As Byte
        Return CByte(((Clr >> &H10) And &HFF))
    End Function
    Public Shared Function ColorG(Clr As Integer) As Byte
        Return CByte(((Clr >> &H8) And &HFF))
    End Function
    Public Shared Function ColorB(Clr As Integer) As Byte
        Return CByte((Clr And &HFF))
    End Function
    Public Shared Function UnsignedToInteger(Value As Long) As Integer
        If Value < 0 Or Value >= 4294967296 Then Throw New OverflowException
        If Value <= 2147483647 Then Return CInt(Value)
        Return CInt(Value - 4294967296)
    End Function
    Public Shared Function MakeArgb(ByVal alpha As Byte, ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte) As Integer
        Return UnsignedToInteger(((((CLng(red) << &H10) Or (CLng(green) << 8)) Or CLng(blue)) Or (CLng(alpha) << &H18)) And &HFFFFFFFF)
    End Function
    Public Shared Function ColorFromName(Clr As String) As Integer
        Dim clrDict As New Dictionary(Of String, Long) From {
            {"Transparent", &HFFFFFF},
           {"AliceBlue", &HFFF0F8FF},
           {"AntiqueWhite", &HFFFAEBD7},
           {"Aqua", &HFF00FFFF},
           {"Aquamarine", &HFF7FFFD4},
           {"Azure", &HFFF0FFFF},
           {"Beige", &HFFF5F5DC},
           {"Bisque", &HFFFFE4C4},
           {"Black", &HFF000000},
           {"BlanchedAlmond", &HFFFFEBCD},
           {"Blue", &HFF0000FF},
           {"BlueViolet", &HFF8A2BE2},
           {"Brown", &HFFA52A2A},
           {"BurlyWood", &HFFDEB887},
           {"CadetBlue", &HFF5F9EA0},
           {"Chartreuse", &HFF7FFF00},
           {"Chocolate", &HFFD2691E},
           {"Coral", &HFFFF7F50},
           {"CornflowerBlue", &HFF6495ED},
           {"Cornsilk", &HFFFFF8DC},
           {"Crimson", &HFFDC143C},
           {"Cyan", &HFF00FFFF},
           {"DarkBlue", &HFF00008B},
           {"DarkCyan", &HFF008B8B},
           {"DarkGoldenrod", &HFFB8860B},
           {"DarkGray", &HFFA9A9A9},
           {"DarkGreen", &HFF006400},
           {"DarkKhaki", &HFFBDB76B},
           {"DarkMagenta", &HFF8B008B},
           {"DarkOliveGreen", &HFF556B2F},
           {"DarkOrange", &HFFFF8C00},
           {"DarkOrchid", &HFF9932CC},
           {"DarkRed", &HFF8B0000},
           {"DarkSalmon", &HFFE9967A},
           {"DarkSeaGreen", &HFF8FBC8B},
           {"DarkSlateBlue", &HFF483D8B},
           {"DarkSlateGray", &HFF2F4F4F},
           {"DarkTurquoise", &HFF00CED1},
           {"DarkViolet", &HFF9400D3},
           {"DeepPink", &HFFFF1493},
           {"DeepSkyBlue", &HFF00BFFF},
           {"DimGray", &HFF696969},
           {"DodgerBlue", &HFF1E90FF},
           {"Firebrick", &HFFB22222},
           {"FloralWhite", &HFFFFFAF0},
           {"ForestGreen", &HFF228B22},
           {"Fuchsia", &HFFFF00FF},
           {"Gainsboro", &HFFDCDCDC},
           {"GhostWhite", &HFFF8F8FF},
           {"Gold", &HFFFFD700},
           {"Goldenrod", &HFFDAA520},
           {"Gray", &HFF808080},
           {"Green", &HFF008000},
           {"GreenYellow", &HFFADFF2F},
           {"Honeydew", &HFFF0FFF0},
           {"HotPink", &HFFFF69B4},
           {"IndianRed", &HFFCD5C5C},
           {"Indigo", &HFF4B0082},
           {"Ivory", &HFFFFFFF0},
           {"Khaki", &HFFF0E68C},
           {"Lavender", &HFFE6E6FA},
           {"LavenderBlush", &HFFFFF0F5},
           {"LawnGreen", &HFF7CFC00},
           {"LemonChiffon", &HFFFFFACD},
           {"LightBlue", &HFFADD8E6},
           {"LightCoral", &HFFF08080},
           {"LightCyan", &HFFE0FFFF},
           {"LightGoldenrodYellow", &HFFFAFAD2},
           {"LightGray", &HFFD3D3D3},
           {"LightGreen", &HFF90EE90},
           {"LightPink", &HFFFFB6C1},
           {"LightSalmon", &HFFFFA07A},
           {"LightSeaGreen", &HFF20B2AA},
           {"LightSkyBlue", &HFF87CEFA},
           {"LightSlateGray", &HFF778899},
           {"LightSteelBlue", &HFFB0C4DE},
           {"LightYellow", &HFFFFFFE0},
           {"Lime", &HFF00FF00},
           {"LimeGreen", &HFF32CD32},
           {"Linen", &HFFFAF0E6},
           {"Magenta", &HFFFF00FF},
           {"Maroon", &HFF800000},
           {"MediumAquamarine", &HFF66CDAA},
           {"MediumBlue", &HFF0000CD},
           {"MediumOrchid", &HFFBA55D3},
           {"MediumPurple", &HFF9370DB},
           {"MediumSeaGreen", &HFF3CB371},
           {"MediumSlateBlue", &HFF7B68EE},
           {"MediumSpringGreen", &HFF00FA9A},
           {"MediumTurquoise", &HFF48D1CC},
           {"MediumVioletRed", &HFFC71585},
           {"MidnightBlue", &HFF191970},
           {"MintCream", &HFFF5FFFA},
           {"MistyRose", &HFFFFE4E1},
           {"Moccasin", &HFFFFE4B5},
           {"NavajoWhite", &HFFFFDEAD},
           {"Navy", &HFF000080},
           {"OldLace", &HFFFDF5E6},
           {"Olive", &HFF808000},
           {"OliveDrab", &HFF6B8E23},
           {"Orange", &HFFFFA500},
           {"OrangeRed", &HFFFF4500},
           {"Orchid", &HFFDA70D6},
           {"PaleGoldenrod", &HFFEEE8AA},
           {"PaleGreen", &HFF98FB98},
           {"PaleTurquoise", &HFFAFEEEE},
           {"PaleVioletRed", &HFFDB7093},
           {"PapayaWhip", &HFFFFEFD5},
           {"PeachPuff", &HFFFFDAB9},
           {"Peru", &HFFCD853F},
           {"Pink", &HFFFFC0CB},
           {"Plum", &HFFDDA0DD},
           {"PowderBlue", &HFFB0E0E6},
           {"Purple", &HFF800080},
           {"Red", &HFFFF0000},
           {"RosyBrown", &HFFBC8F8F},
           {"RoyalBlue", &HFF4169E1},
           {"SaddleBrown", &HFF8B4513},
           {"Salmon", &HFFFA8072},
           {"SandyBrown", &HFFF4A460},
           {"SeaGreen", &HFF2E8B57},
           {"SeaShell", &HFFFFF5EE},
           {"Sienna", &HFFA0522D},
           {"Silver", &HFFC0C0C0},
           {"SkyBlue", &HFF87CEEB},
           {"SlateBlue", &HFF6A5ACD},
           {"SlateGray", &HFF708090},
           {"Snow", &HFFFFFAFA},
           {"SpringGreen", &HFF00FF7F},
           {"SteelBlue", &HFF4682B4},
           {"Tan", &HFFD2B48C},
           {"Teal", &HFF008080},
           {"Thistle", &HFFD8BFD8},
           {"Tomato", &HFFFF6347},
           {"Turquoise", &HFF40E0D0},
           {"Violet", &HFFEE82EE},
           {"Wheat", &HFFF5DEB3},
           {"White", &HFFFFFFFF},
           {"WhiteSmoke", &HFFF5F5F5},
           {"Yellow", &HFFFFFF00},
           {"YellowGreen", &HFF9ACD32}
            }
        Return CInt(clrDict(Clr))
    End Function
    Public Class PrefixComparer
        Implements IComparer(Of String)
        Public Function Compare(ByVal x As String, ByVal y As String) As Integer Implements IComparer(Of String).Compare
            Dim StrLeft As String() = x.Substring(0, CInt(If(x.IndexOf(":") <> -1, x.IndexOf(":"), x.Length))).Split(New Char() {"."c})
            Dim StrRight As String() = y.Substring(0, CInt(If(y.IndexOf(":") <> -1, y.IndexOf(":"), y.Length))).Split(New Char() {"."c})
            If StrLeft.Length = 0 And StrRight.Length = 0 Then Return 0
            If StrLeft.Length = 0 Then Return -1
            If StrRight.Length = 0 Then Return 1
            Dim Check As Integer = String.Compare(StrLeft(0), StrRight(0))
            If Check <> 0 Then Return Check
            If StrLeft.Length = 1 And StrRight.Length = 1 Then Return 0
            If StrLeft.Length = 1 Then Return -1
            If StrRight.Length = 1 Then Return 1
            If StrLeft.Length = 2 And StrRight.Length = 2 Then Return String.Compare(StrLeft(1), StrRight(1))
            If StrLeft.Length = 2 And StrRight.Length = 3 Then Return String.Compare(StrLeft(1), StrRight(2))
            If StrLeft.Length = 3 And StrRight.Length = 2 Then Return String.Compare(StrLeft(2), StrRight(1))
            Check = String.Compare(StrLeft(1), StrRight(1))
            If Check <> 0 Then Return Check
            Return String.Compare(StrLeft(2), StrRight(2))
        End Function
    End Class
    Public Shared Function GetFileLinesByNumberPrefix(Strings() As String, ByVal Prefix As String) As String()
        Dim Index As Integer = Array.BinarySearch(Strings, Prefix, New PrefixComparer)
        If Index < 0 OrElse Index >= Strings.Length OrElse (New PrefixComparer).Compare(Prefix, Strings(Index)) <> 0 Then Return New String() {}
        Dim StartIndex As Integer = Index - 1
        While StartIndex >= 0 _
            AndAlso (New PrefixComparer).Compare(Prefix, Strings(StartIndex)) = 0
            StartIndex -= 1
        End While
        StartIndex += 1
        Index += 1
        While Index < Strings.Length AndAlso
            (New PrefixComparer).Compare(Prefix, Strings(Index)) = 0
            Index += 1
        End While
        Index -= 1
        Dim ReturnStrings(Index - StartIndex) As String
        Array.ConstrainedCopy(Strings, StartIndex, ReturnStrings, 0, Index - StartIndex + 1)
        For Index = 0 To ReturnStrings.Length - 1
            ReturnStrings(Index) = ReturnStrings(Index).Substring(CInt(If(ReturnStrings(Index).IndexOf(":") <> -1, ReturnStrings(Index).IndexOf(":") + 2, 0)))
        Next
        Return ReturnStrings
    End Function
    Public Shared Function GetDigitLength(ByVal Number As Integer) As Integer
        If Number = 0 Then Return 1
        Return CInt(Math.Floor(Math.Log10(Number))) + 1
    End Function
    Public Shared Function ZeroPad(ByVal PadString As String, ByVal ZeroCount As Integer) As String
        Dim RetString As String = New String(New List(Of Char)(Enumerable.Repeat("0"c, ZeroCount)).ToArray()) + PadString
        Return RetString.Substring(RetString.Length - ZeroCount)
    End Function
    Public Shared Function DefaultValue(Value As String, DefValue As String) As String
        If Value Is Nothing Then Return DefValue
        Return Value
    End Function
    Public Shared Function ReadAllLines(Stream As IO.Stream) As String()
        Dim Lines As New List(Of String)
        Dim Reader As New IO.StreamReader(Stream)
        While Not Reader.EndOfStream
            Lines.Add(Reader.ReadLine())
        End While
        Return Lines.ToArray()
    End Function
    Public Shared Function ParseValue(ByVal XMLItemNode As Xml.Linq.XAttribute, ByVal DefaultValue As String) As String
        If XMLItemNode Is Nothing Then
            ParseValue = DefaultValue
        Else
            ParseValue = XMLItemNode.Value
        End If
    End Function
    Public Shared Function GetChildNode(ByVal NodeName As String, ByVal ChildNodes As Xml.Linq.XElement()) As Xml.Linq.XElement
        Dim XMLNode As Xml.Linq.XElement
        For Each XMLNode In ChildNodes
            If XMLNode.Name = NodeName Then
                Return XMLNode
            End If
        Next
        Return Nothing
    End Function
    Public Shared Function GetChildNodes(ByVal NodeName As String, ByVal ChildNodes As Xml.Linq.XElement()) As Xml.Linq.XElement()
        Dim XMLNode As Xml.Linq.XElement
        Dim XMLNodeList As New List(Of Xml.Linq.XElement)
        For Each XMLNode In ChildNodes
            If XMLNode.Name = NodeName Then
                XMLNodeList.Add(XMLNode)
            End If
        Next
        Return XMLNodeList.ToArray()
    End Function
    Public Shared Function GetChildNodeByIndex(ByVal NodeName As String, ByVal IndexName As String, ByVal Index As Integer, ByVal ChildNodes As Xml.Linq.XElement()) As Xml.Linq.XElement
        Dim XMLNode As Xml.Linq.XElement = If(Index >= ChildNodes.Length, Nothing, ChildNodes(Index))
        Dim AttributeNode As Xml.Linq.XAttribute
        If Index - 1 < ChildNodes.Length Then
            XMLNode = ChildNodes(Index - 1)
            If Not XMLNode Is Nothing AndAlso XMLNode.Name = NodeName Then
                AttributeNode = XMLNode.Attribute(IndexName)
                If Not AttributeNode Is Nothing AndAlso CInt(AttributeNode.Value) = Index Then
                    Return XMLNode
                End If
            End If
        End If
        For Each XMLNode In ChildNodes
            If XMLNode.Name = NodeName Then
                AttributeNode = XMLNode.Attribute(IndexName)
                If Not AttributeNode Is Nothing AndAlso CInt(AttributeNode.Value) = Index Then
                    Return XMLNode
                End If
            End If
        Next
        Return Nothing
    End Function
    Public Shared Function GetChildNodeCount(ByVal NodeName As String, ByVal Node As Xml.Linq.XElement) As Integer
        Dim Count As Integer = 0
        For Each Item In Node.Elements
            If Item.Name = NodeName Then Count += 1
        Next
        Return Count
    End Function
    Class CompareNameValueArray
        Implements Collections.IComparer
        'Compares an array of structures with a String and Integer element
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
            Implements Collections.IComparer.Compare
            If CInt(CType(x, Array).GetValue(1)) = CInt(CType(y, Array).GetValue(1)) Then
                Compare = 0
            Else
                Compare = CInt(If(CInt(CType(x, Array).GetValue(1)) > CInt(CType(y, Array).GetValue(1)), 1, -1))
            End If
        End Function
    End Class
End Class
<CLSCompliant(True)>
Public Class ArabicData
    <Runtime.Serialization.DataContract>
    Public Class ArabicCombo
        <Runtime.Serialization.DataMember> Public UnicodeName As String()
        <Runtime.Serialization.DataMember> Public Symbol As Char()
        <Runtime.Serialization.DataMember> Public Shaping() As Char
        Public Function Connecting(Data As ArabicData) As Boolean
            If Not Shaping Is Nothing And Shaping.Length = 1 Then Return Data.ArabicLetters(Data.FindLetterBySymbol(Shaping(0))).Connecting
            Return (Not Shaping Is Nothing AndAlso (Shaping(1) <> Nothing Or Shaping(3) <> Nothing))
        End Function
        Public Function Terminating(Data As ArabicData) As Boolean
            If Not Shaping Is Nothing And Shaping.Length = 1 Then Return Data.ArabicLetters(Data.FindLetterBySymbol(Shaping(0))).Terminating
            Return (Not Shaping Is Nothing AndAlso ((Shaping(0) <> Nothing Or Shaping(1) <> Nothing) And Shaping(2) = Nothing And Shaping(3) = Nothing))
        End Function
    End Class
    Private _ArabicCombos() As ArabicCombo
    <Runtime.Serialization.DataContract>
    Public Class ArabicSymbol
        <Runtime.Serialization.DataMember> Public UnicodeName As String
        <Runtime.Serialization.DataMember> Public Symbol As Char
        <Runtime.Serialization.DataMember> Public Shaping() As Char
        <Runtime.Serialization.DataMember> Public JoiningStyle As String
        <Runtime.Serialization.DataMember> Public CombiningClass As Integer
        Public ReadOnly Property Connecting As Boolean
            Get
                Return JoiningStyle <> "T" AndAlso (JoiningStyle = "final" Or JoiningStyle = "medial" Or JoiningStyle = "C" Or (Not Shaping Is Nothing AndAlso (Shaping(1) <> Nothing Or Shaping(3) <> Nothing)))
            End Get
        End Property
        Public ReadOnly Property Terminating As Boolean
            Get
                Return JoiningStyle <> "T" AndAlso (JoiningStyle = "isolated" Or JoiningStyle = "final" Or JoiningStyle = "U" Or (Not Shaping Is Nothing AndAlso ((Shaping(0) <> Nothing Or Shaping(1) <> Nothing) And Shaping(2) = Nothing And Shaping(3) = Nothing)))
            End Get
        End Property
    End Class
    Private _ArabicLetters() As ArabicSymbol
    Public Async Function LoadArabic() As Task
        Dim LetterBytes As Byte() = Await _PortableMethods.DiskCache.GetCacheItem("ArabicLetters", DateTime.MinValue)
        Dim ComboBytes As Byte() = Await _PortableMethods.DiskCache.GetCacheItem("ArabicCombos", DateTime.MinValue)
        If Not LetterBytes Is Nothing And Not ComboBytes Is Nothing Then
            _ArabicLetters = CType((New Runtime.Serialization.DataContractSerializer(GetType(ArabicData.ArabicSymbol()))).ReadObject(New IO.MemoryStream(LetterBytes)), ArabicData.ArabicSymbol())
            _ArabicCombos = CType((New Runtime.Serialization.DataContractSerializer(GetType(ArabicData.ArabicCombo()))).ReadObject(New IO.MemoryStream(ComboBytes)), ArabicData.ArabicCombo())
            Return
        End If
        Dim CharArr As New List(Of Integer)
        Dim Letters As New List(Of ArabicSymbol)
        Dim Combos As New List(Of ArabicCombo)
        Dim Ranges As List(Of List(Of Integer)) = MakeUniCategory(ALCategories)
        For Count = 0 To Ranges.Count - 1
            Dim Range As List(Of Integer) = Ranges(Count)
            If Range.Count = 1 Then
                CharArr.Add(Range(0))
            Else
                For SubCount = 0 To Range.Count - 1
                    CharArr.Add(Range(SubCount))
                Next
            End If
        Next
        For Count = 0 To CharArr.Count - 1
            If _DecData.ContainsKey(ChrW(CInt(CharArr(Count)))) AndAlso Not _DecData.Item(ChrW(CInt(CharArr(Count)))).Chars Is Nothing AndAlso _DecData.Item(ChrW(CInt(CharArr(Count)))).Chars.Length <> 0 Then
                Dim ComCount As Integer
                For ComCount = 0 To Combos.Count - 1
                    If String.Join(String.Empty, Linq.Enumerable.Select(Combos(ComCount).Symbol, Function(Sym As Char) CStr(Sym))) = String.Join(String.Empty, Linq.Enumerable.Select(_DecData.Item(ChrW(CInt(CharArr(Count)))).Chars, Function(Sym As Char) CStr(Sym))) Then Exit For
                Next
                Dim ArComb As ArabicCombo
                If ComCount = Combos.Count Then
                    ArComb = New ArabicCombo
                    ArComb.Shaping = {Nothing, Nothing, Nothing, Nothing}
                    ArComb.UnicodeName = {Nothing, Nothing, Nothing, Nothing}
                    ArComb.Symbol = _DecData.Item(ChrW(CInt(CharArr(Count)))).Chars
                Else
                    ArComb = Combos(ComCount)
                End If
                Dim Idx As Integer = Array.IndexOf(ShapePositions, _DecData.Item(ChrW(CInt(CharArr(Count)))).JoiningStyle)
                If Idx = -1 Then
                    ArComb.UnicodeName = {_Names.Item(ChrW(CInt(CharArr(Count))))(0)}
                    ArComb.Shaping = {ChrW(CInt(CharArr(Count)))}
                    Dim ArabicLet As New ArabicSymbol
                    ArabicLet.Symbol = ChrW(CInt(CharArr(Count)))
                    ArabicLet.UnicodeName = _Names.Item(ArabicLet.Symbol)(0)
                    ArabicLet.JoiningStyle = _DecData.Item(ArabicLet.Symbol).JoiningStyle
                    ArabicLet.Shaping = _DecData.Item(ArabicLet.Symbol).Shapes
                    Letters.Add(ArabicLet)
                Else
                    ArComb.UnicodeName(Idx) = _Names.Item(ChrW(CInt(CharArr(Count))))(0)
                    ArComb.Shaping(Idx) = ChrW(CInt(CharArr(Count)))
                End If
                If ComCount = Combos.Count Then Combos.Add(ArComb)
            Else
                Dim ArabicLet As New ArabicSymbol
                ArabicLet.Symbol = ChrW(CInt(CharArr(Count)))
                If Array.IndexOf(CombineCategories, _UniClass(ArabicLet.Symbol)) <> -1 Then ArabicLet.JoiningStyle = "T"
                If Array.IndexOf(CausesJoining, ArabicLet.Symbol) <> -1 Then ArabicLet.JoiningStyle = "C"
                If _DecData.ContainsKey(ChrW(CInt(CharArr(Count)))) Then
                    ArabicLet.JoiningStyle = _DecData.Item(ArabicLet.Symbol).JoiningStyle
                    ArabicLet.Shaping = _DecData.Item(ArabicLet.Symbol).Shapes
                End If
                ArabicLet.UnicodeName = _Names.Item(ArabicLet.Symbol)(0)
                Letters.Add(ArabicLet)
            End If
        Next
        CharArr = New List(Of Integer)
        Ranges = MakeUniCategory(WeakCategories)
        For Count = 0 To Ranges.Count - 1
            Dim Range As List(Of Integer) = Ranges(Count)
            If Range.Count = 1 Then
                CharArr.Add(Range(0))
            Else
                For SubCount = 0 To Range.Count - 1
                    CharArr.Add(Range(SubCount))
                Next
            End If
        Next
        For Count = 0 To CharArr.Count - 1
            Dim ArabicLet As New ArabicSymbol
            ArabicLet.Symbol = ChrW(CInt(CharArr(Count)))
            ArabicLet.JoiningStyle = If(Array.IndexOf(CombineCategories, _UniClass(ArabicLet.Symbol)) <> -1, "T", If(Array.IndexOf(CausesJoining, ArabicLet.Symbol) <> -1, "C", "U"))
            ArabicLet.UnicodeName = _Names.Item(ArabicLet.Symbol)(0)
            Letters.Add(ArabicLet)
        Next
        CharArr = New List(Of Integer)
        Ranges = MakeUniCategory(NeutralCategories)
        For Count = 0 To Ranges.Count - 1
            Dim Range As List(Of Integer) = Ranges(Count)
            If Range.Count = 1 Then
                CharArr.Add(Range(0))
            Else
                For SubCount = 0 To Range.Count - 1
                    CharArr.Add(Range(SubCount))
                Next
            End If
        Next
        For Count = 0 To CharArr.Count - 1
            Dim ArabicLet As New ArabicSymbol
            ArabicLet.Symbol = ChrW(CInt(CharArr(Count)))
            ArabicLet.JoiningStyle = If(Array.IndexOf(CombineCategories, _UniClass(ArabicLet.Symbol)) <> -1, "T", If(Array.IndexOf(CausesJoining, ArabicLet.Symbol) <> -1, "C", "U"))
            ArabicLet.UnicodeName = _Names.Item(ArabicLet.Symbol)(0)
            Letters.Add(ArabicLet)
        Next
        _ArabicLetters = Letters.ToArray()
        _ArabicCombos = Combos.ToArray()
        Dim MemStream As New IO.MemoryStream
        Dim Ser As New Runtime.Serialization.DataContractSerializer(GetType(ArabicData.ArabicSymbol()))
        Ser.WriteObject(MemStream, _ArabicLetters)
        Await _PortableMethods.DiskCache.CacheItem("ArabicLetters", DateTime.Now, MemStream.ToArray())
        MemStream.Dispose()
        MemStream = New IO.MemoryStream
        Ser = New Runtime.Serialization.DataContractSerializer(GetType(ArabicData.ArabicCombo()))
        Ser.WriteObject(MemStream, _ArabicCombos)
        Await _PortableMethods.DiskCache.CacheItem("ArabicCombos", DateTime.Now, MemStream.ToArray())
        MemStream.Dispose()
    End Function
    Public ReadOnly Property ArabicCombos() As ArabicCombo()
        Get
            Return _ArabicCombos
        End Get
    End Property
    Public ReadOnly Property ArabicLetters() As ArabicSymbol()
        Get
            Return _ArabicLetters
        End Get
    End Property
    Public Function GetUnicodeName(Character As Char) As String
        Dim Str As String = _PortableMethods.Settings.GetUName(Character)
        If Str <> String.Empty Then Return Str
        If FindLetterBySymbol(Character) = -1 Then Return String.Empty
        Dim Res As String = _PortableMethods.LoadResourceString("unicode_" + ArabicLetters(FindLetterBySymbol(Character)).UnicodeName)
        If Res.Length <> 0 Then Return Res
        Return ArabicLetters(FindLetterBySymbol(Character)).UnicodeName
    End Function
    Private CamelCaseRegEx As System.Text.RegularExpressions.Regex
    Public Function ToCamelCase(Str As String) As String
        Return CamelCaseRegEx.Replace(Str, Function(CamCase As System.Text.RegularExpressions.Match) String.Concat(CamCase.Groups(1).Value, CamCase.Groups(2).Value.ToLowerInvariant()))
    End Function
    Public Function IsTerminating(Str As String, Index As Integer) As Boolean
        Dim bIsEnd = True 'default to non-connecting end
        'should probably check for any non-arabic letters also
        For CharCount As Integer = Index + 1 To Str.Length - 1
            Dim Idx As Integer = FindLetterBySymbol(Str(CharCount))
            If Idx = -1 OrElse ArabicLetters()(Idx).JoiningStyle <> "T" Then
                bIsEnd = Idx = -1 OrElse Not ArabicLetters()(Idx).Connecting
                Exit For
            End If
        Next
        Return bIsEnd
    End Function
    Public Function IsLastConnecting(Str As String, Index As Integer) As Boolean
        Dim bLastConnects = False 'default to non-connecting beginning 
        For CharCount As Integer = Index - 1 To 0 Step -1
            Dim Idx As Integer = FindLetterBySymbol(Str(CharCount))
            If Idx <> -1 AndAlso ArabicLetters(Idx).JoiningStyle <> "T" Then
                bLastConnects = Idx <> -1 AndAlso Not ArabicLetters(Idx).Terminating
                Exit For
            End If
        Next
        Return bLastConnects
    End Function
    Public Shared Function GetShapeIndex(bConnects As Boolean, bLastConnects As Boolean, bIsEnd As Boolean) As Integer
        If Not bLastConnects And (Not bConnects Or bConnects And bIsEnd) Then
            Return 0
        ElseIf bLastConnects And (Not bConnects Or bConnects And bIsEnd) Then
            Return 1
        ElseIf Not bLastConnects And bConnects And Not bIsEnd Then
            Return 2
        ElseIf bLastConnects And bConnects And Not bIsEnd Then
            Return 3
        End If
        Return -1
    End Function
    Public Function GetShapeIndexFromString(Str As String, Index As Integer, Length As Integer) As Integer
        'ignore all transparent characters
        'isolated - non-connecting + (non-connecting letter | connecting letter + end)
        'final - connecting + (non-connecting letter | connecting letter + end)
        'initial - non-connecting + connecting letter + not end
        'medial - connecting + connecting letter + not end
        Dim bIsEnd As Boolean = IsTerminating(Str, Index + Length - 1)
        Dim Idx As Integer = FindLetterBySymbol(Str.Chars(Index + Length - 1))
        Dim bConnects As Boolean = Not ArabicLetters()(Idx).Terminating
        Dim bLastConnects As Boolean = ArabicLetters()(Idx).Connecting And IsLastConnecting(Str, Index)
        Return GetShapeIndex(bConnects, bLastConnects, bIsEnd)
    End Function
    Public Function TransformChars(Str As String) As String
        For Count As Integer = 0 To ArabicCombos.Length - 1
            If ArabicCombos(Count).Shaping.Length = 1 Then
                Str = Str.Replace(String.Join(String.Empty, Linq.Enumerable.Select(ArabicCombos(Count).Symbol, Function(Sym As Char) CStr(Sym))), ArabicCombos(Count).Shaping(0))
            End If
        Next
        Return Str
    End Function
    Public Class LigatureInfo
        Public Ligature As String
        Public Indexes() As Integer
    End Class
    Public Function GetFormsRange(BeginIndex As Char, EndIndex As Char) As Char()
        Dim Forms As New List(Of Char)
        For Count As Integer = 0 To ArabicCombos.Length - 1
            If Not ArabicCombos(Count).Shaping Is Nothing Then
                For Each Shape As Char In ArabicCombos(Count).Shaping
                    If Shape >= BeginIndex AndAlso Shape <= EndIndex Then Forms.Add(Shape)
                Next
            End If
        Next
        For Count As Integer = 0 To ArabicLetters().Length - 1
            If Not ArabicLetters(Count).Shaping Is Nothing Then
                For Each Shape As Char In ArabicLetters()(Count).Shaping
                    If Shape >= BeginIndex AndAlso Shape <= EndIndex Then Forms.Add(Shape)
                Next
            End If
        Next
        Return Forms.ToArray()
    End Function
    Private _PresentationForms() As Char
    Private _PresentationFormsA() As Char
    Private _PresentationFormsB() As Char
    Public ReadOnly Property GetPresentationForms() As Char()
        Get
            Return _PresentationForms
        End Get
    End Property
    Public ReadOnly Property GetPresentationFormsA() As Char()
        Get
            Return _PresentationFormsA
        End Get
    End Property
    Public ReadOnly Property GetPresentationFormsB() As Char()
        Get
            Return _PresentationFormsB
        End Get
    End Property
    Public Function CheckLigatureMatch(Str As String, CurPos As Integer, ByRef Positions As Integer()) As Integer
        'if first is 2 diacritics or letter + diacritic
        'letter + diacritic done only unless a space present as 2 diacritics could be nexted in required ligature which would be skipped
        'must check space with 2 diacritics first, second check will already capture space with diacritic
        If Str.Length > 2 AndAlso FindLetterBySymbol(Str(1)) <> -1 AndAlso ArabicLetters(FindLetterBySymbol(Str(1))).JoiningStyle = "T" AndAlso ArabicLetters(FindLetterBySymbol(Str(2))).JoiningStyle = "T" AndAlso (LigatureLookups.ContainsKey(Str.Substring(0, 3)) Or LigatureLookups.ContainsKey(Str(0) + Str(2) + Str(1))) Then
            Positions = {CurPos, CurPos + 1, CurPos + 2}
            Return LigatureLookups.Item(If(LigatureLookups.ContainsKey(Str.Substring(0, 3)), Str.Substring(0, 3), Str(0) + Str(2) + Str(1)))
        ElseIf Str.Length > 1 AndAlso FindLetterBySymbol(Str(1)) <> -1 AndAlso ArabicLetters(FindLetterBySymbol(Str(1))).JoiningStyle = "T" AndAlso LigatureLookups.ContainsKey(Str.Substring(0, 2)) Then
            Positions = {CurPos, CurPos + 1}
            Return LigatureLookups.Item(Str.Substring(0, 2))
        End If
        If FindLetterBySymbol(Str(0)) <> -1 AndAlso ArabicLetters(FindLetterBySymbol(Str(0))).JoiningStyle <> "T" Then
            'only 3 letters or 2 letters has possible parsing, or several 4 and a multiword 8 and 18
            Dim StrCount As Integer = 0
            Positions = {CurPos + StrCount}
            For Count = 1 To 18
                StrCount += 1
                While StrCount <> Str.Length AndAlso FindLetterBySymbol(Str(StrCount)) <> -1 AndAlso ArabicLetters(FindLetterBySymbol(Str(StrCount))).JoiningStyle = "T"
                    StrCount += 1
                End While
                If StrCount = Str.Length Then Exit For
                ReDim Preserve Positions(Count)
                Positions(Count) = CurPos + StrCount
            Next
            If Positions.Length = 1 Then Positions = {}
            While Positions.Length <> 0
                If LigatureLookups.ContainsKey(String.Join(String.Empty, Linq.Enumerable.Select(Positions, Function(Pos As Integer) CStr(Str(Pos - CurPos))))) Then
                    Return LigatureLookups.Item(String.Join(String.Empty, Linq.Enumerable.Select(Positions, Function(Pos As Integer) CStr(Str(Pos - CurPos)))))
                End If
                ReDim Preserve Positions(Positions.Length - 2)
            End While
        End If
        'if first is diacritic or letter
        'check space diacritic first
        If Str.Length > 1 AndAlso FindLetterBySymbol(Str(1)) <> -1 AndAlso ArabicLetters(FindLetterBySymbol(Str(1))).JoiningStyle = "T" AndAlso (LigatureLookups.ContainsKey(" " + Str.Substring(0, 2)) Or LigatureLookups.ContainsKey(" " + Str(1) + Str(0))) Then
            Positions = {CurPos, CurPos + 1}
            Return LigatureLookups.Item(" " + If(LigatureLookups.ContainsKey(" " + Str.Substring(0, 2)), Str.Substring(0, 2), Str(1) + Str(0)))
        ElseIf Str.Length > 2 AndAlso FindLetterBySymbol(Str(1)) <> -1 AndAlso ArabicLetters(FindLetterBySymbol(Str(1))).JoiningStyle = "T" AndAlso ArabicLetters(FindLetterBySymbol(Str(2))).JoiningStyle = "T" AndAlso (LigatureLookups.ContainsKey(Str(0) + Str(2))) Then
            'Tatweel with a Hamza Above followed by diacritic
            Positions = {CurPos, CurPos + 2}
            Return LigatureLookups.Item(Str(0) + Str(2))
        ElseIf LigatureLookups.ContainsKey(Str.Substring(0, 1)) Then
            Positions = {CurPos}
            Return LigatureLookups.Item(Str.Substring(0, 1))
        ElseIf LigatureLookups.ContainsKey(" " + Str.Substring(0, 1)) Then
            Positions = {CurPos}
            Return LigatureLookups.Item(" " + Str.Substring(0, 1))
        End If
        Return -1
    End Function
    Private _LigatureCombos() As ArabicCombo
    Public ReadOnly Property LigatureCombos As ArabicCombo()
        Get
            Return _LigatureCombos
        End Get
    End Property
    Private _LigatureShapes As Dictionary(Of Char, Integer)
    Public ReadOnly Property LigatureShapes As Dictionary(Of Char, Integer)
        Get
            Return _LigatureShapes
        End Get
    End Property
    Private _LigatureLookups As Dictionary(Of String, Integer)
    Public ReadOnly Property LigatureLookups As Dictionary(Of String, Integer)
        Get
            Return _LigatureLookups
        End Get
    End Property
    Public Function GetLigatures(Str As String, Dir As Boolean, SupportedForms As Char()) As LigatureInfo()
        Dim Count As Integer
        Dim SubCount As Integer
        Dim Ligatures As New List(Of LigatureInfo)
        Dim Combos As ArabicCombo() = LigatureCombos
        'Division selection between Presentation A and B forms can be done here though wasl and gunnah need consideration
        Count = 0
        While Count <> Str.Length
            If Dir Then
                If LigatureShapes.ContainsKey(Str.Chars(Count)) Then
                    'ZWJ and ZWNJ could be used to preserve deliberately improper shaped Arabic or other strategies beyond just default shaping
                    Ligatures.Add(New LigatureInfo With {.Ligature = Combos(LigatureShapes.Item(Str.Chars(Count))).Symbol, .Indexes = {Count}})
                End If
            Else
                Dim Indexes As Integer() = Nothing
                SubCount = CheckLigatureMatch(Str.Substring(Count), Count, Indexes)
                'transform ligatures are not processed here
                If SubCount <> -1 AndAlso Combos(SubCount).Shaping <> Nothing AndAlso Combos(SubCount).Shaping.Length <> 1 Then
                    Dim Index As Integer = Linq.Enumerable.TakeWhile(Combos(SubCount).Symbol, Function(Ch As Char) Not (Ch = " " Or FindLetterBySymbol(Ch) <> -1 AndAlso (ArabicLetters(FindLetterBySymbol(Ch)).JoiningStyle = "T" Or ArabicLetters(FindLetterBySymbol(Ch)).JoiningStyle = "C"))).Count()
                    'diacritics always use isolated form sitting on a space which is actually optional
                    Dim Shape As Integer = If(Index = 0, If(FindLetterBySymbol(Combos(SubCount).Symbol(Index)) <> -1 AndAlso ArabicLetters(FindLetterBySymbol(Combos(SubCount).Symbol(Index))).JoiningStyle = "C", 3, 0), GetShapeIndexFromString(Str, Count, Indexes(Indexes.Length - 1) - Count + 1 - If(Index = Combos(SubCount).Symbol.Length, 0, Index)))
                    If Combos(SubCount).Shaping(Shape) <> ChrW(0) AndAlso Array.IndexOf(SupportedForms, Combos(SubCount).Shaping(Shape)) <> -1 Then
                        Ligatures.Add(New LigatureInfo With {.Ligature = Combos(SubCount).Shaping(Shape), .Indexes = Indexes})
                        'Ligatures can surround other ligatures which represents significant challenge
                    End If
                End If
            End If
            Count += 1
            While Linq.Enumerable.TakeWhile(Ligatures, Function(Lig As LigatureInfo) Array.IndexOf(Lig.Indexes, Count) = -1).Count() <> Ligatures.Count
                Count += 1
            End While
        End While
        Return Ligatures.ToArray()
    End Function
    Public Function ConvertLigatures(Str As String, Dir As Boolean, SupportedForms As Char()) As String
        Dim Ligatures() As LigatureInfo = GetLigatures(Str, Dir, SupportedForms)
        For Count = Ligatures.Length - 1 To 0 Step -1
            For Index = 0 To Ligatures(Count).Indexes.Length - 1
                Str = Str.Remove(Ligatures(Count).Indexes(Index), 1).Insert(Ligatures(Count).Indexes(0), Ligatures(Count).Ligature)
            Next
        Next
        Return Str
    End Function
    Private _ArabicLetterMap As Dictionary(Of Char, Integer)
    Public ReadOnly Property ArabicLetterMap() As Dictionary(Of Char, Integer)
        Get
            Return _ArabicLetterMap
        End Get
    End Property
    Public Function FindLetterBySymbol(Symbol As Char) As Integer
        Return If(ArabicLetterMap.ContainsKey(Symbol), ArabicLetterMap.Item(Symbol), -1)
    End Function
    Public Const Space As Char = ChrW(&H20)
    Public Const ExclamationMark As Char = ChrW(&H21)
    Public Const QuotationMark As Char = ChrW(&H22)
    Public Const Comma As Char = ChrW(&H2C)
    Public Const FullStop As Char = ChrW(&H2E)
    Public Const HyphenMinus As Char = ChrW(&H2D)
    Public Const Colon As Char = ChrW(&H3A)
    Public Const LeftParenthesis As Char = ChrW(&H5B)
    Public Const RightParenthesis As Char = ChrW(&H5D)
    Public Const LeftSquareBracket As Char = ChrW(&H5B)
    Public Const RightSquareBracket As Char = ChrW(&H5D)
    Public Const LeftCurlyBracket As Char = ChrW(&H7B)
    Public Const RightCurlyBracket As Char = ChrW(&H7D)
    Public Const NoBreakSpace As Char = ChrW(&HA0)
    Public Const LeftPointingDoubleAngleQuotationMark As Char = ChrW(&HAB)
    Public Const RightPointingDoubleAngleQuotationMark As Char = ChrW(&HBB)
    Public Const ArabicComma As Char = ChrW(&H60C)
    Public Const ArabicSignSallallahouAlayheWassallam As Char = ChrW(&H610)
    Public Const ArabicLetterHamza As Char = ChrW(&H621)
    Public Const ArabicLetterAlefWithMaddaAbove As Char = ChrW(&H622)
    Public Const ArabicLetterAlefWithHamzaAbove As Char = ChrW(&H623)
    Public Const ArabicLetterWawWithHamzaAbove As Char = ChrW(&H624)
    Public Const ArabicLetterAlefWithHamzaBelow As Char = ChrW(&H625)
    Public Const ArabicLetterYehWithHamzaAbove As Char = ChrW(&H626)
    Public Const ArabicLetterAlef As Char = ChrW(&H627)
    Public Const ArabicLetterBeh As Char = ChrW(&H628)
    Public Const ArabicLetterTehMarbuta As Char = ChrW(&H629)
    Public Const ArabicLetterTeh As Char = ChrW(&H62A)
    Public Const ArabicLetterTheh As Char = ChrW(&H62B)
    Public Const ArabicLetterJeem As Char = ChrW(&H62C)
    Public Const ArabicLetterHah As Char = ChrW(&H62D)
    Public Const ArabicLetterKhah As Char = ChrW(&H62E)
    Public Const ArabicLetterDal As Char = ChrW(&H62F)
    Public Const ArabicLetterThal As Char = ChrW(&H630)
    Public Const ArabicLetterReh As Char = ChrW(&H631)
    Public Const ArabicLetterZain As Char = ChrW(&H632)
    Public Const ArabicLetterSeen As Char = ChrW(&H633)
    Public Const ArabicLetterSheen As Char = ChrW(&H634)
    Public Const ArabicLetterSad As Char = ChrW(&H635)
    Public Const ArabicLetterDad As Char = ChrW(&H636)
    Public Const ArabicLetterTah As Char = ChrW(&H637)
    Public Const ArabicLetterZah As Char = ChrW(&H638)
    Public Const ArabicLetterAin As Char = ChrW(&H639)
    Public Const ArabicLetterGhain As Char = ChrW(&H63A)
    Public Const ArabicTatweel As Char = ChrW(&H640)
    Public Const ArabicLetterFeh As Char = ChrW(&H641)
    Public Const ArabicLetterQaf As Char = ChrW(&H642)
    Public Const ArabicLetterKaf As Char = ChrW(&H643)
    Public Const ArabicLetterLam As Char = ChrW(&H644)
    Public Const ArabicLetterMeem As Char = ChrW(&H645)
    Public Const ArabicLetterNoon As Char = ChrW(&H646)
    Public Const ArabicLetterHeh As Char = ChrW(&H647)
    Public Const ArabicLetterWaw As Char = ChrW(&H648)
    Public Const ArabicLetterAlefMaksura As Char = ChrW(&H649)
    Public Const ArabicLetterYeh As Char = ChrW(&H64A)

    Public Const ArabicFathatan As Char = ChrW(&H64B)
    Public Const ArabicDammatan As Char = ChrW(&H64C)
    Public Const ArabicKasratan As Char = ChrW(&H64D)
    Public Const ArabicFatha As Char = ChrW(&H64E)
    Public Const ArabicDamma As Char = ChrW(&H64F)
    Public Const ArabicKasra As Char = ChrW(&H650)
    Public Const ArabicShadda As Char = ChrW(&H651)
    Public Const ArabicSukun As Char = ChrW(&H652)
    Public Const ArabicMaddahAbove As Char = ChrW(&H653)
    Public Const ArabicHamzaAbove As Char = ChrW(&H654)
    Public Const ArabicHamzaBelow As Char = ChrW(&H655)
    Public Const ArabicVowelSignDotBelow As Char = ChrW(&H65C)
    Public Const Bullet As Char = ChrW(&H2022)
    Public Const ArabicLetterSuperscriptAlef As Char = ChrW(&H670)
    Public Const ArabicLetterAlefWasla As Char = ChrW(&H671)
    Public Const ArabicSmallHighLigatureSadWithLamWithAlefMaksura As Char = ChrW(&H6D6)
    Public Const ArabicSmallHighLigatureQafWithLamWithAlefMaksura As Char = ChrW(&H6D7)
    Public Const ArabicSmallHighMeemInitialForm As Char = ChrW(&H6D8)
    Public Const ArabicSmallHighLamAlef As Char = ChrW(&H6D9)
    Public Const ArabicSmallHighJeem As Char = ChrW(&H6DA)
    Public Const ArabicSmallHighThreeDots As Char = ChrW(&H6DB)
    Public Const ArabicSmallHighSeen As Char = ChrW(&H6DC)
    Public Const ArabicEndOfAyah As Char = ChrW(&H6DD)
    Public Const ArabicStartOfRubElHizb As Char = ChrW(&H6DE)
    Public Const ArabicSmallHighRoundedZero As Char = ChrW(&H6DF)
    Public Const ArabicSmallHighUprightRectangularZero As Char = ChrW(&H6E0)
    Public Const ArabicSmallHighMeemIsolatedForm As Char = ChrW(&H6E2)
    Public Const ArabicSmallLowSeen As Char = ChrW(&H6E3)
    Public Const ArabicSmallWaw As Char = ChrW(&H6E5)
    Public Const ArabicSmallYeh As Char = ChrW(&H6E6)
    Public Const ArabicSmallHighNoon As Char = ChrW(&H6E8)
    Public Const ArabicPlaceOfSajdah As Char = ChrW(&H6E9)
    Public Const ArabicEmptyCentreLowStop As Char = ChrW(&H6EA)
    Public Const ArabicEmptyCentreHighStop As Char = ChrW(&H6EB)
    Public Const ArabicRoundedHighStopWithFilledCentre As Char = ChrW(&H6EC)
    Public Const ArabicSmallLowMeem As Char = ChrW(&H6ED)
    Public Const ArabicSemicolon As Char = ChrW(&H61B)
    Public Const ArabicLetterMark As Char = ChrW(&H61C)
    Public Const ArabicQuestionMark As Char = ChrW(&H61F)
    Public Const ArabicLetterPeh As Char = ChrW(&H67E)
    Public Const ArabicLetterTcheh As Char = ChrW(&H686)
    Public Const ArabicLetterVeh As Char = ChrW(&H6A4)
    Public Const ArabicLetterGaf As Char = ChrW(&H6AF)
    Public Const ArabicLetterNoonGhunna As Char = ChrW(&H6BA)
    Public Const ZeroWidthSpace As Char = ChrW(&H200B)
    Public Const ZeroWidthNonJoiner As Char = ChrW(&H200C)
    Public Const ZeroWidthJoiner As Char = ChrW(&H200D)
    Public Const LeftToRightMark As Char = ChrW(&H200E)
    Public Const RightToLeftMark As Char = ChrW(&H200F)
    Public Const PopDirectionalFormatting As Char = ChrW(&H202C)
    Public Const LeftToRightEmbedding As Char = ChrW(&H202A)
    Public Const RightToLeftEmbedding As Char = ChrW(&H202B)
    Public Const LeftToRightOverride As Char = ChrW(&H202D)
    Public Const RightToLeftOverride As Char = ChrW(&H202E)
    Public Const NarrowNoBreakSpace As Char = ChrW(&H202F)
    Public Const LeftToRightIsolate As Char = ChrW(&H2066)
    Public Const RightToLeftIsolate As Char = ChrW(&H2067)
    Public Const FirstStrongIsolate As Char = ChrW(&H2068)
    Public Const PopDirectionalIsolate As Char = ChrW(&H2069)
    Public Const DottedCircle As Char = ChrW(&H25CC)
    Public Const OrnateLeftParenthesis As Char = ChrW(&HFD3E)
    Public Const OrnateRightParenthesis As Char = ChrW(&HFD3F)
    'http://www.unicode.org/Public/7.0.0/ucd/UnicodeData.txt
    Public Shared LTRCategories As String() = New String() {"L"}
    Public Shared RTLCategories As String() = New String() {"R", "AL"}
    Public Shared ALCategories As String() = New String() {"AL"}
    Public Shared CombineCategories As String() = New String() {"Mn", "Me", "Cf"}
    Public Shared NeutralCategories As String() = New String() {"B", "S", "WS", "ON"}
    Public Shared WeakCategories As String() = New String() {"EN", "ES", "ET", "AN", "CS", "NSM", "BN"}
    Public Shared ExplicitCategories As String() = New String() {"LRE", "LRO", "RLE", "RLO", "PDF", "LRI", "RLI", "FSI", "PDI"}
    Public Shared CausesJoining As Char() = New Char() {ArabicTatweel, ZeroWidthJoiner}
    Private _PortableMethods As PortableMethods
    Public Sub New(NewPortableMethods As PortableMethods)
        _PortableMethods = NewPortableMethods
    End Sub
    Public Async Function Init() As Task
        'Await GetJoiningData()
        Await GetDecompositionCombiningCatData()
        Await LoadArabic()
        _ArabicLetterMap = New Dictionary(Of Char, Integer)
        For Index = 0 To ArabicLetters.Length - 1
            If ArabicLetters(Index).Symbol <> ChrW(0) Then
                _ArabicLetterMap.Add(ArabicLetters(Index).Symbol, Index)
            End If
        Next
        ReDim _LigatureCombos(ArabicLetters.Length + ArabicCombos.Length - 1)
        ArabicCombos.CopyTo(_LigatureCombos, 0)
        For Count = 0 To ArabicLetters.Length - 1
            'do not need to transfer UnicodeName as it is not used here
            _LigatureCombos(ArabicCombos.Length + Count) = New ArabicCombo
            _LigatureCombos(ArabicCombos.Length + Count).Symbol = {ArabicLetters(Count).Symbol}
            _LigatureCombos(ArabicCombos.Length + Count).Shaping = ArabicLetters(Count).Shaping
        Next
        Array.Sort(_LigatureCombos, Function(Com1 As ArabicCombo, Com2 As ArabicCombo) If(Com1.Symbol.Length = Com2.Symbol.Length, String.Join(String.Empty, Linq.Enumerable.Select(Com1.Symbol, Function(Sym As Char) CStr(Sym))).CompareTo(String.Join(String.Empty, Linq.Enumerable.Select(Com2.Symbol, Function(Sym As Char) CStr(Sym)))), If(Com1.Symbol.Length > Com2.Symbol.Length, -1, 1)))
        _LigatureLookups = New Dictionary(Of String, Integer)
        For Count = 0 To _LigatureCombos.Length - 1
            'If there is only an isolated form then the combos which come before letters would take precedence
            If Not _LigatureCombos(Count).Shaping Is Nothing And Not _LigatureLookups.ContainsKey(String.Join(String.Empty, Linq.Enumerable.Select(_LigatureCombos(Count).Symbol, Function(Sym As Char) CStr(Sym)))) Then
                _LigatureLookups.Add(String.Join(String.Empty, Linq.Enumerable.Select(_LigatureCombos(Count).Symbol, Function(Sym As Char) CStr(Sym))), Count)
            End If
        Next
        _LigatureShapes = New Dictionary(Of Char, Integer)
        For Count As Integer = 0 To _LigatureCombos.Length - 1
            If Not _LigatureCombos(Count).Shaping Is Nothing Then
                For SubCount As Integer = 0 To _LigatureCombos(Count).Shaping.Length - 1
                    If Not _LigatureShapes.ContainsKey(_LigatureCombos(Count).Shaping(SubCount)) Then _LigatureShapes.Add(_LigatureCombos(Count).Shaping(SubCount), Count)
                Next
            End If
        Next
        _PresentationFormsA = GetFormsRange(ChrW(&HFB50), ChrW(&HFDFF))
        _PresentationFormsB = GetFormsRange(ChrW(&HFE70), ChrW(&HFEFF))
        Dim Forms As New List(Of Char)
        Forms.AddRange(GetPresentationFormsA())
        Forms.AddRange(GetPresentationFormsB())
        _PresentationForms = Forms.ToArray()
        CamelCaseRegEx = New System.Text.RegularExpressions.Regex("(?:([A-Z])([A-Z]+)?)?(-|\s+|$)")
    End Function
    Public Async Function GetJoiningData() As Task(Of Dictionary(Of Char, String))
        Dim Strs As String() = Await _PortableMethods.ReadAllLines(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", "ArabicShaping.txt")))
        Dim Joiners As New Dictionary(Of Char, String)
        Dim Vals As String()
        For Count = 0 To Strs.Length - 1
            If Strs(Count).Length <> 0 AndAlso Strs(Count)(0) <> "#" Then
                Vals = Strs(Count).Split(";"c)
                'C Join_Causing on Tatweel and ZeroWidthJoiner could be considered as Dual_Joining
                'General Category Mn, Me, or Cf are T Transparent and all others are U Non_Joining
                Joiners.Add(ChrW(Integer.Parse(Vals(0), Globalization.NumberStyles.AllowHexSpecifier)), Vals(3))
            End If
        Next
        Return Joiners
    End Function
    Class DecData
        Public JoiningStyle As String
        Public Chars As Char()
        Public Shapes As Char()
    End Class
    Public Shared ShapePositions As String() = {"isolated", "final", "initial", "medial"}
    Private Shared _CombPos As Dictionary(Of Char, Integer)
    Private Shared _UniClass As Dictionary(Of Char, String)
    Private Shared _DecData As Dictionary(Of Char, DecData)
    Private Shared _Ranges As Dictionary(Of String, List(Of List(Of Integer)))
    Private Shared _Names As Dictionary(Of Char, String())
    Public Shared Sub Split(str As String, separator As Char, ByRef sepList As List(Of Integer), ByRef strs As String())
        sepList.Clear()
        Dim length As Integer = str.Length
        Dim j As Integer = &H0
        Do While ((j < str.Length) AndAlso (sepList.Count < length))
            j = str.IndexOf(separator, j)
            If j = -1 Then Exit Do
            sepList.Add(j)
            j += 1
        Loop
        If (sepList.Count = 0) Then strs(0) = str
        Dim startIndex As Integer = &H0
        Dim index As Integer = &H0
        Dim i As Integer = &H0
        Do While ((i < sepList.Count) AndAlso (startIndex < str.Length))
            strs(index) = str.Substring(startIndex, (sepList(i) - startIndex))
            index += 1
            startIndex = (sepList(i) + &H1)
            i += 1
        Loop
        If ((startIndex < str.Length) AndAlso (sepList.Count >= &H0)) Then
            strs(index) = str.Substring(startIndex)
        End If
        If (index = sepList.Count) Then
            strs(index) = String.Empty
        End If
    End Sub
    Public Async Function GetDecompositionCombiningCatData() As Task
        Dim Strs As String() = Await _PortableMethods.ReadAllLines(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", "UnicodeData.txt")))
        _CombPos = New Dictionary(Of Char, Integer)
        _UniClass = New Dictionary(Of Char, String)
        _Ranges = New Dictionary(Of String, List(Of List(Of Integer)))
        _DecData = New Dictionary(Of Char, DecData)
        _Names = New Dictionary(Of Char, String())
        Dim sepList As New List(Of Integer)
        Dim Vals(15) As String
        For Count = 0 To Strs.Length - 1
            Split(Strs(Count), ";"c, sepList, Vals)
            'All symbol categories not needed
            If (Vals(2)(0) = "S" And Vals(4) <> "ON") Or Integer.Parse(Vals(0), Globalization.NumberStyles.AllowHexSpecifier) >= &H10000 Then Continue For
            Dim Ch As Char = ChrW(Integer.Parse(Vals(0), Globalization.NumberStyles.AllowHexSpecifier))
            _UniClass.Add(Ch, Vals(2))
            If Vals(5) <> "" Then
                Dim CombData As String() = Vals(5).Split(" "c)
                If Not _DecData.ContainsKey(Ch) Then _DecData.Add(Ch, New DecData With {.Shapes = New Char() {Nothing, Nothing, Nothing, Nothing}})
                Dim Data As DecData = _DecData(Ch)
                If CombData(0).StartsWith("<") And CombData(0).EndsWith(">") Then
                    Data.JoiningStyle = CombData(0).Trim("<"c, ">"c)
                    ReDim Data.Chars(CombData.Length - 2)
                    For SubCount = 0 To CombData.Length - 2
                        Data.Chars(SubCount) = ChrW(Integer.Parse(CombData(SubCount + 1), Globalization.NumberStyles.AllowHexSpecifier))
                    Next
                    _DecData(Ch) = Data
                    If CombData.Length = 2 Then
                        If Not _DecData.ContainsKey(Data.Chars(0)) Then _DecData.Add(Data.Chars(0), New DecData With {.Shapes = New Char() {Nothing, Nothing, Nothing, Nothing}})
                        Dim ShapeData As DecData = _DecData(Data.Chars(0))
                        If Array.IndexOf(ShapePositions, Data.JoiningStyle) <> -1 Then ShapeData.Shapes(Array.IndexOf(ShapePositions, Data.JoiningStyle)) = Ch
                    End If
                Else
                    Data.Chars = Linq.Enumerable.Select(CombData, Function(Dat As String) ChrW(If(Integer.Parse(Dat, Globalization.NumberStyles.AllowHexSpecifier) >= &H10000, 0, Integer.Parse(Dat, Globalization.NumberStyles.AllowHexSpecifier)))).ToArray()
                    _DecData(Ch) = Data
                End If
            End If
            If Vals(3) <> "" Then
                _CombPos.Add(Ch, Integer.Parse(Vals(3), Globalization.NumberStyles.Integer))
            End If
            If Vals(10) <> "" Then
                _Names.Add(Ch, {Vals(1), Vals(10)})
            Else
                _Names.Add(Ch, {Vals(1)})
            End If
            Dim NewRangeMatch As Integer = Integer.Parse(Vals(0), Globalization.NumberStyles.AllowHexSpecifier)
            If Not _Ranges.ContainsKey(Vals(4)) Then _Ranges.Add(Vals(4), New List(Of List(Of Integer)))
            If _Ranges(Vals(4)).Count <> 0 AndAlso CInt(_Ranges(Vals(4))(_Ranges(Vals(4)).Count - 1)(_Ranges(Vals(4))(_Ranges(Vals(4)).Count - 1).Count - 1)) + 1 = NewRangeMatch Then
                _Ranges(Vals(4))(_Ranges(Vals(4)).Count - 1).Add(NewRangeMatch)
            Else
                _Ranges(Vals(4)).Add(New List(Of Integer) From {NewRangeMatch})
            End If
        Next
    End Function
    Public Function MakeUniCategory(Cats As String()) As List(Of List(Of Integer))
        Dim Ranges As New List(Of List(Of Integer))
        For Count = 0 To Cats.Length - 1
            If _Ranges.ContainsKey(Cats(Count)) Then
                Ranges.AddRange(_Ranges(Cats(Count)))
            End If
        Next
        Return Ranges
    End Function
    Public Function IsDiacritic(Ch As Char) As Boolean
        Return ArabicLetters(FindLetterBySymbol(Ch)).JoiningStyle = "T"
    End Function
    Public Function FixStartingCombiningSymbol(Str As String) As String
        Return If((FindLetterBySymbol(Str.Chars(0)) <> -1 AndAlso ArabicLetters(FindLetterBySymbol(Str.Chars(0))).JoiningStyle = "T") Or Str.Length = 1, LeftToRightOverride + Str + PopDirectionalFormatting, Str)
    End Function
    Public Shared IsArabicExp As New System.Text.RegularExpressions.Regex("[\p{IsArabic}\p{IsArabicPresentationForms-A}\p{IsArabicPresentationForms-B}]")
    Public Shared Function MakeUniRegEx(Input As String) As String
        Return String.Join(String.Empty, Linq.Enumerable.Select(Of Char, String)(Input.ToCharArray(), Function(Ch As Char) If(IsArabicExp.Match(Ch).Success, "\u" + AscW(Ch).ToString("X4"), If(Ch = "."c, "\" + Ch, Ch))))
    End Function
    Public Shared Function MakeRegMultiEx(Input As String()) As String
        Return String.Join("|"c, Input)
    End Function
    Public Enum TranslitScheme
        None = 0
        Literal = 1
        RuleBased = 2
        LearningMode = 3
    End Enum
End Class
Public Class RenderArray
    Enum RenderTypes
        eHeaderLeft 'these 3 should be text and combined display class, they are not that distinct
        eHeaderCenter
        eHeaderRight
        eText
        eInteractive
    End Enum
    Enum RenderDisplayClass
        eNested
        eArabic
        eTransliteration
        eTransliterationRTL
        eLTR
        eRTL
        eContinueStop
        eRanking
        eList
        eTag
        eLink
        ePassThru
        eReference
    End Enum
    Class RenderText
        Public DisplayClass As RenderDisplayClass
        Public Clr As Integer 'Color
        Public Text As Object
        Public Font As String
        Sub New(ByVal NewDisplayClass As RenderDisplayClass, ByVal NewText As Object)
            DisplayClass = NewDisplayClass
            Text = NewText
            Clr = Utility.MakeArgb(&HFF, 0, 0, 0) 'Color.Black 'default
            Font = String.Empty
        End Sub
    End Class
    Class RenderItem
        Public Type As RenderTypes
        Public TextItems() As RenderText
        Sub New(ByVal NewType As RenderTypes, ByVal NewTextItems() As RenderText)
            Type = NewType
            TextItems = NewTextItems
        End Sub
    End Class
    Public Sub New(ID As String)
        _ID = ID
    End Sub
    Public Items As New Collections.Generic.List(Of RenderItem)
    Public _ID As String
End Class