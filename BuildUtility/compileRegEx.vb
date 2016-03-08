Imports HostPageUtility
Imports XMLRender

Module compileRegEx
    Public Class WindowsSettings
        Implements PortableSettings
        Public ReadOnly Property CacheDirectory As String Implements PortableSettings.CacheDirectory
            Get
                Return IO.Directory.GetCurrentDirectory()
            End Get
        End Property
        Public ReadOnly Property Resources As KeyValuePair(Of String, String())() Implements PortableSettings.Resources
            Get
                Return (New List(Of KeyValuePair(Of String, String()))(Linq.Enumerable.Select("XMLRender=Acct,lang,unicode;IslamResources=Hadith,IslamInfo,IslamSource".Split(";"c), Function(Str As String) New KeyValuePair(Of String, String())(Str.Split("="c)(0), Str.Split("="c)(1).Split(","c))))).ToArray()
            End Get
        End Property
        Public ReadOnly Property FuncLibs As String() Implements PortableSettings.FuncLibs
            Get
                Return {"IslamMetadata"}
            End Get
        End Property
        Public Function GetTemplatePath() As String Implements PortableSettings.GetTemplatePath
            Return GetFilePath("metadata\IslamSource.xml")
        End Function
        Public Function GetFilePath(ByVal Path As String) As String Implements PortableSettings.GetFilePath
            Return "..\..\..\" + Path
        End Function
        <Runtime.InteropServices.DllImport("getuname.dll", EntryPoint:="GetUName")>
        Friend Shared Function GetUName(ByVal wCharCode As UShort, <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.LPWStr)> ByVal lpbuf As System.Text.StringBuilder) As Integer
        End Function
        Public Function GetUName(Character As Char) As String Implements PortableSettings.GetUName
            Dim Str As New System.Text.StringBuilder(512)
            Try
                GetUName(CUShort(AscW(Character)), Str)
            Catch e As System.DllNotFoundException
            End Try
            Return Str.ToString()
        End Function
        Public Function GetResourceString(baseName As String, resourceKey As String) As String Implements PortableSettings.GetResourceString
            Return New System.Resources.ResourceManager(baseName + ".Resources", Reflection.Assembly.Load(baseName)).GetString(resourceKey, Threading.Thread.CurrentThread.CurrentUICulture)
        End Function
    End Class
    Sub Main()
        PortableMethods.FileIO = New WindowsWebFileIO
        PortableMethods.Settings = New WindowsSettings
        Dim RegExs As New List(Of System.Text.RegularExpressions.RegexCompilationInfo)
        For Count As Integer = 0 To IslamMetadata.CachedData.IslamData.RuleSets.Count - 1
            For SubCount As Integer = 0 To IslamMetadata.CachedData.IslamData.RuleSets(Count).Rules.Count - 1
                RegExs.Add(New System.Text.RegularExpressions.RegexCompilationInfo(IslamMetadata.CachedData.TranslateRegEx(IslamMetadata.CachedData.IslamData.RuleSets(Count).Rules(SubCount).Match, True), Text.RegularExpressions.RegexOptions.None, IslamMetadata.CachedData.IslamData.RuleSets(Count).Name + "_" + CStr(SubCount), String.Empty, True))
            Next
        Next
        For Count As Integer = 0 To IslamMetadata.CachedData.IslamData.MetaRules.Count - 1
            For SubCount As Integer = 0 To IslamMetadata.CachedData.IslamData.MetaRules(Count).Rules.Count - 1
                RegExs.Add(New System.Text.RegularExpressions.RegexCompilationInfo(IslamMetadata.CachedData.TranslateRegEx(IslamMetadata.CachedData.IslamData.MetaRules(Count).Rules(SubCount).Match, True), Text.RegularExpressions.RegexOptions.None, IslamMetadata.CachedData.IslamData.MetaRules(Count).Name + "_" + CStr(SubCount), String.Empty, True))
            Next
        Next
        System.Text.RegularExpressions.Regex.CompileToAssembly(RegExs.ToArray(), New Reflection.AssemblyName("IslamMetadataRegEx, Version=1.0.0.1001, Culture=neutral, PublicKeyToken=null"))
    End Sub

End Module
