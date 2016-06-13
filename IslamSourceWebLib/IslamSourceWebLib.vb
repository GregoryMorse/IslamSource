Option Explicit On
Option Strict On
Imports IslamMetadata
Imports XMLRender
Imports HostPageUtility
Imports System.Drawing
Imports System.Web
Imports System.Web.UI
Public Class InitClass
    Implements Utility.IInitClass
    Private _PortableMethods As PortableMethods
    Private ChData As CachedData
    Private Arb As Arabic
    Private ArbWeb As ArabicWeb
    Private ArbData As ArabicData
    Private ArbDataWeb As ArabicDataWeb
    Private TR As TanzilReader
    Private DocBuild As DocBuilder
    Private ArbFont As ArabicFont
    Private ChDataWeb As CachedDataWeb
    Private DBWeb As DocBuilderWeb
    Private PhrWeb As PhrasesWeb
    Private Qz As Quiz
    Private TRWeb As TanzilReaderWeb
    Private HR As HadithReader
    Private HRWeb As HadithReaderWeb
    Private UWeb As UtilityWeb
    Private SD As SiteDatabase
    Private ISDatabase As IslamSiteDatabase
    Private ISDatabaseLookup As IslamSiteDatabaseLookup
    Private IMInit As IslamMetadata.InitClass
    Public Sub New(NewPortableMethods As PortableMethods, NewArbData As ArabicData, NewUWeb As UtilityWeb, NewSD As SiteDatabase)
        _PortableMethods = NewPortableMethods
        ArbData = NewArbData
        UWeb = NewUWeb
        SD = NewSD
    End Sub
    Public Async Function Init() As Task Implements Utility.IInitClass.Init
        ArbDataWeb = New ArabicDataWeb(ArbData)
        IMInit = New IslamMetadata.InitClass(_PortableMethods, ArbData)
        Await IMInit.Init()
        Arb = IMInit.Arb
        ChData = IMInit.ChData
        TR = IMInit.TR
        DocBuild = IMInit.DocBuild
        ArbWeb = New ArabicWeb(_PortableMethods, ChData, Arb, ArbData, ArbDataWeb)
        ArbFont = New ArabicFont(_PortableMethods, ChData)
        ChDataWeb = New CachedDataWeb(ChData)
        TRWeb = New TanzilReaderWeb(_PortableMethods, ChData, Arb, ArbWeb, TR)
        PhrWeb = New PhrasesWeb(ArbWeb, ChData, TR)
        DBWeb = New DocBuilderWeb(_PortableMethods, ArbWeb, ChData, Arb, ArbData, TR, DocBuild, PhrWeb, TRWeb)
        PhrWeb.Init(DBWeb)
        Qz = New Quiz(Arb, ArbWeb, ChData, ArbDataWeb)
        ISDatabase = New IslamSiteDatabase(SD)
        ISDatabaseLookup = New IslamSiteDatabaseLookup(UWeb, ISDatabase)
        HR = New HadithReader(_PortableMethods, Arb, ChData)
        HRWeb = New HadithReaderWeb(_PortableMethods, Arb, ChData, ArbWeb, HR, UWeb, ISDatabaseLookup)
    End Function
    Public Function LookupObject(ClassName As String) As Object Implements Utility.IInitClass.LookupObject
        If ClassName = "ArabicWeb" Then
            Return ArbWeb
        ElseIf ClassName = "ArabicFont" Then
            Return ArbFont
        ElseIf ClassName = "CachedDataWeb" Then
            Return ChDataWeb
        ElseIf ClassName = "TanzilReaderWeb" Then
            Return TRWeb
        ElseIf ClassName = "DocBuilderWeb" Then
            Return DBWeb
        ElseIf ClassName = "PhrasesWeb" Then
            Return PhrWeb
        ElseIf ClassName = "Quiz" Then
            Return Qz
        ElseIf ClassName = "IslamSiteDatabase" Then
            Return ISDatabase
        ElseIf ClassName = "IslamSiteDatabaseLookup" Then
            Return ISDatabaseLookup
        ElseIf ClassName = "HadithReaderWeb" Then
            Return HRWeb
        Else
            Return Nothing
        End If
    End Function
    Public Function GetDependency() As Nullable(Of KeyValuePair(Of String, Utility.IInitClass)) Implements Utility.IInitClass.GetDependency
        Return New KeyValuePair(Of String, Utility.IInitClass)("IslamMetadata", IMInit)
    End Function
End Class
Public Class PrayerTimeWeb
    Public Shared Function GetMonthName(ByVal Item As PageLoader.TextItem) As String
        Return PrayerTime.GetMonthName(Item.Name)
    End Function
    Public Shared Function GetCalendar(ByVal Item As PageLoader.TextItem) As Array()
        Return PrayerTime.GetCalendar(Item.Name)
    End Function
    Public Shared Function GetPrayerTimes(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim Strings As String() = Geolocation.GetGeoData(Context)
        If Strings.Length <> 11 OrElse Strings(0) = "ERROR" Then Return New Array() {}
        Return PrayerTime.GetPrayerTimes(Strings, Geolocation.GetElevationData(Strings(8), Strings(9)))
    End Function
    Public Shared Function GetQiblaDirection(ByVal Item As PageLoader.TextItem, Context As HttpContext) As String
        Return PrayerTime.GetQiblaDirection(Geolocation.GetGeoData(Context))
    End Function
End Class
Public Class ArabicWeb
    Private _PortableMethods As PortableMethods
    Private ChData As CachedData
    Private Arb As Arabic
    Private ArbData As ArabicData
    Private ArbDataWeb As ArabicDataWeb
    Public Sub New(NewPortableMethods As PortableMethods, NewChData As CachedData, NewArb As Arabic, NewArbData As ArabicData, NewArbDataWeb As ArabicDataWeb)
        _PortableMethods = NewPortableMethods
        ChData = NewChData
        Arb = NewArb
        ArbData = NewArbData
        ArbDataWeb = NewArbDataWeb
        TransliterateGenJS = {
            FindLetterBySymbolJS,
            "function isLetterDiacritic(index) { return (" + String.Join("||", Linq.Enumerable.Select(ChData.RecitationLettersDiacritics, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Convert.ToString(AscW(C.Chars(0)), 16))) + "); }",
            "function isSpecialSymbol(index) { return (" + String.Join("||", Linq.Enumerable.Select(ChData.RecitationSpecialSymbols, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Convert.ToString(AscW(C.Chars(0)), 16))) + "); }",
            "function isCombiningSymbol(index) { return (" + String.Join("||", Linq.Enumerable.Select(ChData.RecitationCombiningSymbols, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Convert.ToString(AscW(C.Chars(0)), 16))) + "); }",
            "var arabicMultis = " + UtilityWeb.MakeJSArray(ChData.ArabicMultis) + ";",
            "var arabicSpecials = " + UtilityWeb.MakeJSArray(ChData.ArabicSpecialLetters) + ";",
            "function getSchemeSpecialFromMatch(str, bExp) { if (bExp) { for (var count = 0; count < arabicSpecials.length; count++) { re = new RegExp(arabicSpecials[count], ''); if (re.exec(str) !== null) return count; } } else { return arabicSpecials.indexOf(str); } return -1; }",
            "function generateDefaultStops(str) { var re = new RegExp('(^\s*|\s+)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + "[\p{Nd}]{1,3}|" + ChData.OptionalPattern + ")(?=\s*$|\s+)', 'g'); var arr, defstops = [], dottoggle = false; while ((arr = re.exec(str)) !== null) { if (arr[2] === String.fromCharCode(0x6D6) && (arr[2] !== String.fromCharCode(0x6DB) || dottoggle)) defstops.push(arr.index + arr[1].length); if (arr[2] === String.fromCharCode(0x6DB)) dottoggle = !dottoggle; } }",
            "function doTransliterate(sVal, direction, conversion) { var iCount, iSubCount, sOutVal = ''; if (conversion === 0) return sVal; if (direction && (conversion % 2) === 0) return transliterateWithRules(sVal, Math.floor((conversion - 2) / 2) + 2, generateDefaultStops(sVal), false); for (iCount = 0; iCount < sVal.length; iCount++) { if (sVal.charAt(iCount) === '\\' || direction && (sVal.charAt(iCount) === String.fromCharCode(0x60C) || sVal.charAt(iCount) === String.fromCharCode(0x61B) || sVal.charAt(iCount) === String.fromCharCode(0x61F))) { if (!direction) iCount++; if (sVal.charAt(iCount) === (direction ? String.fromCharCode(0x60C) : ',')) { sOutVal += (direction ? '\\,' : String.fromCharCode(0x60C)); } else if (sVal.charAt(iCount) === (direction ? String.fromCharCode(0x61B) : ';')) { sOutVal += (direction ? '\\;' : String.fromCharCode(0x61B)); } else if (sVal.charAt(iCount) === (direction ? String.fromCharCode(0x61F) : '?')) { sOutVal += (direction ? '\\?' : String.fromCharCode(0x61F)); } else { sOutVal += String.fromCharCode(0x202A) + sVal.charAt(iCount) + String.fromCharCode(0x202C); } } else { if (getSchemeSpecialFromMatch(sVal.slice(iCount), false) !== -1) { sOutVal += translitSchemes[(Math.floor((conversion - 2) / 2) + 2).toString()].special[getSchemeSpecialFromMatch(sVal.slice(iCount), false)]; re = new RegExp(arabicSpecials[getSchemeSpecialFromMatch(sVal.slice(iCount), false)], ''); iCount += re.exec(sVal.slice(iCount))[0].length - 1; } else if (sVal.length - iCount > 1 && arabicMultis.indexOf(sVal.slice(iCount, 2)) !== -1) { sOutVal += translitSchemes[(Math.floor((conversion - 2) / 2) + 2).toString()].multi[arabicMultis.indexOf(sVal.slice(iCount, 2))]; iCount++; } else { for (iSubCount = 0; iSubCount < arabicLetters.length; iSubCount++) { if (direction ? sVal.charCodeAt(iCount) === parseInt(arabicLetters[iSubCount].Symbol, 10) : sVal.charAt(iCount) === unescape(translitSchemes[(Math.floor((conversion - 2) / 2) + 2).toString()].standard[String.fromCharCode(arabicLetters[iSubCount].Symbol)])) { sOutVal += (direction ? (translitSchemes[(Math.floor((conversion - 2) / 2) + 2).toString()].standard[String.fromCharCode(arabicLetters[iSubCount].Symbol)] ? translitSchemes[(Math.floor((conversion - 2) / 2) + 2).toString()].standard[String.fromCharCode(arabicLetters[iSubCount].Symbol)] : '') : (((isCombiningSymbol(iSubCount) && (iSubCount === 0 || findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1)) === -1 || !isLetterDiacritic(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))) && !isSpecialSymbol(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))))) ? String.fromCharCode(0x202B) : '') + String.fromCharCode(arabicLetters[iSubCount].Symbol) + ((isCombiningSymbol(iSubCount) && (iSubCount === 0 || findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1)) === -1 || !isLetterDiacritic(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))) && !isSpecialSymbol(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))))) ? String.fromCharCode(0x202C) : ''))); break; } } if (iSubCount === arabicLetters.length && sVal.charCodeAt(iCount) !== 0x200E && sVal.charCodeAt(iCount) !== 0x200F && !IsExplicit(sVal.charCodeAt(iCount))) sOutVal += ((direction && conversion === 1 && sVal.charAt(iCount) !== '\n') ? '\\' : '') + sVal.charAt(iCount); } } } return unescape(sOutVal); }"
        }
        IsDiacriticJS = "function isDiacritic(index) { return (" + String.Join("||", Linq.Enumerable.Select(ChData.RecitationDiacritics, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Convert.ToString(AscW(C.Chars(0)), 16))) + "); }"
        PlainTransliterateGenJS = {FindLetterBySymbolJS, IsDiacriticJS,
            "function isWhitespace(index) { return (" + String.Join("||", Linq.Enumerable.Select(ChData.WhitespaceSymbols, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Convert.ToString(AscW(C.Chars(0)), 16))) + "); }",
            "function isPunctuation(index) { return (" + String.Join("||", Linq.Enumerable.Select(ChData.PunctuationSymbols, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Convert.ToString(AscW(C.Chars(0)), 16))) + "); }",
            "function isStop(index) { return (" + String.Join("||", Linq.Enumerable.Select(ChData.ArabicStopLetters, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Convert.ToString(AscW(C.Chars(0)), 16))) + "); }",
            "function applyColorRules(sVal) {}",
            "function changeScript(sVal, scriptType) {}",
            "var arabicAlphabet = " + UtilityWeb.MakeJSArray(ChData.ArabicAlphabet) + ";",
            "var arabicLettersInOrder = " + UtilityWeb.MakeJSArray(ChData.ArabicLettersInOrder) + ";",
            "var arabicLeadingGutterals = " + UtilityWeb.MakeJSArray(ChData.ArabicLeadingGutterals) + ";",
            "function getSchemeGutteralFromString(str, scheme, leading) { if (arabicLeadingGutterals.indexOf(str) !== -1) { return translitSchemes[scheme].gutteral[arabicLeadingGutterals.indexOf(str) + (leading ? arabicLeadingGutterals.length : 0)]; } return ''; }",
            "function arabicLetterSpelling(sVal, bQuranic) { var count, index, output = ''; for (count = 0; count < sVal.length; count++) { index = findLetterBySymbol(sVal.charCodeAt(count)); if (index !== -1 && isLetter(index)) { if (output !== '' && !bQuranic) output += ' '; var idx = arabicLettersInOrder.indexOf(String.fromCharCode(parseInt(arabicLetters[index].Symbol, 10))); output += arabicAlphabet[idx].slice(0, -1) + ((arabicAlphabet[idx][arabicAlphabet[idx].length - 1] == 'n') ? '' : 'o'); } else if (index !== -1 && arabicLetters[index].Symbol === 0x653) { output += sVal.charCodeAt(count); } } return doTransliterate(output, false, 1); }",
            "function schemeHasValue(str, scheme) { for (var k in translitSchemes[scheme]) { if (translitSchemes[scheme].hasOwnProperty(k) && str === translitSchemes[scheme][k]) return true; } return false; }",
            "String.prototype.format = function() { var formatted = this; for (var i = 0; i < arguments.length; i++) { formatted = formatted.replace(new RegExp('\\{'+i+'\\}', 'gi'), arguments[i]); } return formatted; };",
            "RegExp.matchResult = function(subexp, offset, str, matches) { return subexp.replace(/\$(\$|&|`|\'|[0-9]+)/g, function(m, p) { if (p === '$') return '$'; if (p === '`') return str.slice(0, offset); if (p === '\'') return str.slice(offset + matches[0].length); if (p === '&' || parseInt(p, 10) <= 0 || parseInt(p, 10) >= matches.length) return matches[0]; return (matches[parseInt(p, 10)] === undefined) ? '' : matches[parseInt(p, 10)]; }); };",
            "var ruleFunctions = [function(str, scheme, learningMode) { return [str.toUpperCase()]; }, function(str, scheme, learningMode) { return [transliterateWithRules(doTransliterate(arabicWordFromNumber(parseInt(doTransliterate(str, true, 1), 10), true, false, false), false, 1), scheme, null, learningMode)]; }, function(str, scheme, learningMode) { return [transliterateWithRules(arabicLetterSpelling(str, true), scheme, null, learningMode)]; }, function(str, scheme, learningMode) { return [translitSchemes[scheme.toString()].standard[str]]; }, function(str, scheme, learningMode) { return [translitSchemes[scheme.toString()].multi[arabicMultis.indexOf(str)]]; }, function(str, scheme, learningMode) { return [" + UtilityWeb.MakeJSArray(ChData.ArabicFathaDammaKasra) + "[" + UtilityWeb.MakeJSArray(ChData.ArabicTanweens) + ".indexOf(str)], '" + ArabicData.ArabicLetterNoon + "']; }, function (str, scheme, learningMode) { return [getSchemeGutteralFromString(str.slice(0, -1), scheme, true) + str[str.length - 1]]; }, function(str, scheme, learningMode) { return [str[0] + getSchemeGutteralFromString(str.slice(1), scheme, false)]; }, function(str, scheme, learningMode) { return [schemeHasValue(translitSchemes[scheme.toString()].standard[str[0]] + translitSchemes[scheme.toString()].standard[str[1]]) ? str[0] + '-' + str[1] : str]; }, function(str, scheme, learningMode) { return learningMode ? [str, ''] : ['', str]; }];",
            "function isLetter(index) { return (" + String.Join("||", Linq.Enumerable.Select(ChData.RecitationLetters, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Convert.ToString(AscW(C.Chars(0)), 16))) + "); }",
            "function isSymbol(index) { return (" + String.Join("||", Linq.Enumerable.Select(Arb.GetRecitationSymbols(ArabicData.TranslitScheme.None, String.Empty), Function(A As Array) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Convert.ToString(AscW(ArbData.ArabicLetters(CInt(A.GetValue(1))).Symbol), 16))) + "); }",
            "var uthmaniMinimalScript = " + UtilityWeb.MakeJSArray(New String() {UtilityWeb.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "negativematch", "ruleFunc"},
                                    New List(Of Object())(Linq.Enumerable.Select(ChData.RuleTranslations("UthmaniMinimalScript"), Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {UtilityWeb.MakeJSString(Convert.Name), UtilityWeb.MakeJSString(UtilityWeb.EscapeJS(Convert.Match)), UtilityWeb.MakeJSString(Convert.Evaluator), UtilityWeb.MakeJSString(Convert.NegativeMatch), Convert.RuleFunc})).ToArray(), True)}, True) + ";",
            "var simpleEnhancedScript = " + UtilityWeb.MakeJSArray(New String() {UtilityWeb.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "negativematch", "ruleFunc"},
                                    New List(Of Object())(Linq.Enumerable.Select(ChData.RuleTranslations("SimpleEnhancedScript"), Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {UtilityWeb.MakeJSString(Convert.Name), UtilityWeb.MakeJSString(UtilityWeb.EscapeJS(Convert.Match)), UtilityWeb.MakeJSString(Convert.Evaluator), UtilityWeb.MakeJSString(Convert.NegativeMatch), Convert.RuleFunc})).ToArray(), True)}, True) + ";",
            "var simpleScriptBase = " + UtilityWeb.MakeJSArray(New String() {UtilityWeb.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "negativematch", "ruleFunc"},
                                    New List(Of Object())(Linq.Enumerable.Select(ChData.RuleTranslations("SimpleScriptBase"), Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {UtilityWeb.MakeJSString(Convert.Name), UtilityWeb.MakeJSString(UtilityWeb.EscapeJS(Convert.Match)), UtilityWeb.MakeJSString(Convert.Evaluator), UtilityWeb.MakeJSString(Convert.NegativeMatch), Convert.RuleFunc})).ToArray(), True)}, True) + ";",
            "var simpleScript = " + UtilityWeb.MakeJSArray(New String() {UtilityWeb.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "negativematch", "ruleFunc"},
                                    New List(Of Object())(Linq.Enumerable.Select(ChData.RuleTranslations("SimpleScript"), Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {UtilityWeb.MakeJSString(Convert.Name), UtilityWeb.MakeJSString(UtilityWeb.EscapeJS(Convert.Match)), UtilityWeb.MakeJSString(Convert.Evaluator), UtilityWeb.MakeJSString(Convert.NegativeMatch), Convert.RuleFunc})).ToArray(), True)}, True) + ";",
            "var simpleMinimalScript = " + UtilityWeb.MakeJSArray(New String() {UtilityWeb.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "negativematch", "ruleFunc"},
                                    New List(Of Object())(Linq.Enumerable.Select(ChData.RuleTranslations("SimpleMinimalScript"), Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {UtilityWeb.MakeJSString(Convert.Name), UtilityWeb.MakeJSString(UtilityWeb.EscapeJS(Convert.Match)), UtilityWeb.MakeJSString(Convert.Evaluator), UtilityWeb.MakeJSString(Convert.NegativeMatch), Convert.RuleFunc})).ToArray(), True)}, True) + ";",
            "var simpleCleanScript = " + UtilityWeb.MakeJSArray(New String() {UtilityWeb.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "negativematch", "ruleFunc"},
                                    New List(Of Object())(Linq.Enumerable.Select(ChData.RuleTranslations("SimpleCleanScript"), Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {UtilityWeb.MakeJSString(Convert.Name), UtilityWeb.MakeJSString(UtilityWeb.EscapeJS(Convert.Match)), UtilityWeb.MakeJSString(Convert.Evaluator), UtilityWeb.MakeJSString(Convert.NegativeMatch), Convert.RuleFunc})).ToArray(), True)}, True) + ";",
            "var errorCheckRules = " + UtilityWeb.MakeJSArray(New String() {UtilityWeb.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "negativematch", "ruleFunc"},
                                    New List(Of Object())(Linq.Enumerable.Select(ChData.RuleTranslations("ErrorCheck"), Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {UtilityWeb.MakeJSString(Convert.Name), UtilityWeb.MakeJSString(UtilityWeb.EscapeJS(Convert.Match)), UtilityWeb.MakeJSString(Convert.Evaluator), UtilityWeb.MakeJSString(Convert.NegativeMatch), Convert.RuleFunc})).ToArray(), True)}, True) + ";",
            "var coloringSpelledOutRules = " + UtilityWeb.MakeJSArray(New String() {UtilityWeb.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "ruleFunc"},
                                    New List(Of Object())(Linq.Enumerable.Select(Of IslamData.ColorRuleCategory.ColorRule, Object())(ChData.IslamData.ColorRuleSets(1).ColorRules, Function(Convert As IslamData.ColorRuleCategory.ColorRule) New Object() {UtilityWeb.MakeJSString(Convert.Name), UtilityWeb.MakeJSString(UtilityWeb.EscapeJS(Convert._Match)), UtilityWeb.MakeJSString(Convert.Evaluator), Convert._MetaRuleFunc})).ToArray(), True)}, True) + ";",
            "var romanizationRules = " + UtilityWeb.MakeJSArray(New String() {UtilityWeb.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "ruleFunc"},
                                    New List(Of Object())(Linq.Enumerable.Select(ChData.RuleTranslations("RomanizationRules"), Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {UtilityWeb.MakeJSString(Convert.Name), UtilityWeb.MakeJSString(UtilityWeb.EscapeJS(Convert.Match)), UtilityWeb.MakeJSString(Convert.Evaluator), Convert.RuleFunc})).ToArray(), True)}, True) + ";",
            "var coloringRules = " + UtilityWeb.MakeJSArray(New String() {UtilityWeb.MakeJSIndexedObject(New String() {"rule", "match", "color"},
                                    New List(Of Object())(Linq.Enumerable.Select(Of IslamData.ColorRuleCategory.ColorRule, String())(ChData.IslamData.ColorRuleSets(0).ColorRules, Function(Convert As IslamData.ColorRuleCategory.ColorRule) New String() {Convert.Name, UtilityWeb.EscapeJS(Convert._Match), System.Drawing.ColorTranslator.ToHtml(Color.FromArgb(Convert.Color))})).ToArray(), False)}, True) + ";",
            "var rulesOfRecitationRegEx = " + UtilityWeb.MakeJSArray(New String() {UtilityWeb.MakeJSIndexedObject(New String() {"rule", "match", "evaluator"},
                                    New List(Of Object())(Linq.Enumerable.Select(Of IslamData.RuleMetaSet.RuleMetadataTranslation, Object())(ChData.RuleMetas("UthmaniQuran").Rules, Function(Convert As IslamData.RuleMetaSet.RuleMetadataTranslation) New Object() {UtilityWeb.MakeJSString(Convert.Name), UtilityWeb.MakeJSString(UtilityWeb.EscapeJS(Convert.Match)), If(Convert.Evaluator Is Nothing, Nothing, UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(Convert.Evaluator, Function(Str As IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs()) (String.Join("|"c, Linq.Enumerable.Select(Str, Function(S) S.RuleName + If(S.Args Is Nothing OrElse S.Args.Length = 0, "", "(" + String.Join(",", Linq.Enumerable.Select(S.Args, Function(Arg) String.Join(" ", Arg))) + ")")))))).ToArray()))})).ToArray(), True)}, True) + ";",
            "var allowZeroLength = " + UtilityWeb.MakeJSArray(Arabic.AllowZeroLength) + ";",
            "function ruleMetadataComparer(a, b) { return (a.index === b.index) ? (b.length === a.length ? b.origOrder - a.origOrder : b.length - a.length) : b.index - a.index; }",
            "function replaceMetadata(sVal, metadataRule, scheme, learningMode) { var count, elimParen = function(s) { return s.replace(/\(.*\)/, ''); }; for (count = 0; count < coloringSpelledOutRules.length; count++) { var index, match = null; for (index = 0; index < coloringSpelledOutRules[count].match.split('|').length; index++) { if (metadataRule.type.split('|').map(elimParen).indexOf(coloringSpelledOutRules[count].match.split('|')[index]) !== -1) { match = coloringSpelledOutRules[count].match.split('|')[index]; break; } } if (match !== null) { var str = coloringSpelledOutRules[count].evaluator.format(sVal.substr(metadataRule.index, metadataRule.length)); if (coloringSpelledOutRules[count].ruleFunc !== 0) { var args = ruleFunctions[coloringSpelledOutRules[count].ruleFunc - 1](str, scheme, learningMode); if (args.length === 1) { str = args[0]; } else { var metaArgs = metadataRule.type.match(/\((.*)\)/)[1].split(','); str = ''; for (index = 0; index < args.length; index++) { if (args[index] && (learningMode || coloringSpelledOutRules[count].ruleFunc !== " + CStr(Arabic.MetaRuleFuncs.eLearningMode) + " || index !== 0)) str += replaceMetadata(args[index], {index: 0, length: args[index].length, type: metaArgs[index].replace(' ', '|'), origOrder: index}, scheme, learningMode); } } } sVal = sVal.substr(0, metadataRule.index) + str + sVal.substr(metadataRule.index + metadataRule.length); } } return sVal; }",
            "function joinContig(sVal, preString, postString) { var index = preString.lastIndexOf(' '); if (index !== -1 && preString.length - 2 === index) index = preString.lastIndexOf(' ', index - 1); if (index !== -1) preString = preString.substring(index + 1); if (preString !== '') preString += ' ' + String.fromCharCode(0x6DD) + ' '; index = postString.indexOf(' '); if (index === 1) index = preString.indexOf(' ', index + 1); if (index !== -1) postString = postString.substring(0, index); if (postString !== '') postString = ' ' + String.fromCharCode(0x6DD) + ' ' + postString; return preString + sVal + postString; }",
            "function unjoinContig(sVal, preString, postString) { var index = sVal.indexOf(String.fromCharCode(0x6DD)); if (preString !== '' && index !== -1) sVal = sVal.substring(index + 1 +  1); index = sVal.lastIndexOf(String.fromCharCode(0x6DD)); if (postString !== '' && index !== -1) sVal = sVal.substring(0, index - 1); return sVal; }",
            "function transliterateContigWithRules(sVal, preString, postString, scheme, optionalStops) { return unjoinContig(transliterateWithRules(JoinContig(sVal, preString, postString), scheme, optionalStops, false), preString, postString); }",
            "function transliterateWithRules(sVal, scheme, optionalStops, learningMode) { var count, index, arr, re, metadataList = [], replaceFunc = function(f, e) { return function() { return f(RegExp.matchResult(e, arguments[arguments.length - 2], arguments[arguments.length - 1], Array.prototype.slice.call(arguments).slice(0, -2)), scheme)[0]; }; }; for (count = 0; count < errorCheckRules.length; count++) { re = new RegExp(errorCheckRules[count].match, 'g'); while ((arr = re.exec(sVal)) !== null) { if (!errorCheckRules[count].negativematch || RegExp.matchResult(errorCheckRules[count].negativematch, arr.index, sVal, arr) === '') { console.log(errorCheckRules[count].rule + ': ' + doTransliterate(sVal.substr(0, arr.index), true, 1) + '<!-- -->' + doTransliterate(sVal.substr(arr.index), true, 1)); } } } for (count = 0; count < rulesOfRecitationRegEx.length; count++) { if (rulesOfRecitationRegEx[count].evaluator !== null) { var subcount, lindex; re = new RegExp(rulesOfRecitationRegEx[count].match, 'g'); while ((arr = re.exec(sVal)) !== null) { lindex = arr.index; for (subcount = 0; subcount < rulesOfRecitationRegEx[count].evaluator.length; subcount++) { if (rulesOfRecitationRegEx[count].evaluator[subcount].split('|').indexOf('optionalstop') !== -1 && (optionalStops === null && arr[subcount + 1] === String.fromCharCode(0x6D6) || (arr[subcount + 1] !== undefined && lindex !== 0 && lindex !== sVal.length) || (optionalStops !== null && arr[subcount + 1] && optionalStops.indexOf(lindex) === -1)) || rulesOfRecitationRegEx[count].evaluator[subcount].split('|').indexOf('optionalnotstop') !== -1 && (optionalStops === null && arr[subcount + 1] !== String.fromCharCode(0x6D6) && ((arr[subcount + 1] !== undefined && arr[subcount + 1].length === 0) && (lindex === 0 || lindex === sVal.length)) || optionalStops !== null && arr[subcount + 1] && optionalStops.indexOf(lindex) !== -1)) break; } if (subcount !== rulesOfRecitationRegEx[count].evaluator.length) continue; for (subcount = 0; subcount < rulesOfRecitationRegEx[count].evaluator.length; subcount++) { if (rulesOfRecitationRegEx[count].evaluator[subcount] !== null && rulesOfRecitationRegEx[count].evaluator[subcount] !== '' && (arr[subcount + 1] && arr[subcount + 1].length !== 0 || allowZeroLength.indexOf(rulesOfRecitationRegEx[count].evaluator[subcount]) !== -1)) { metadataList.push({index: lindex, length: arr[subcount + 1] ? arr[subcount + 1].length : 0, type: rulesOfRecitationRegEx[count].evaluator[subcount], origOrder: subcount}); } lindex += (arr[subcount + 1] ? arr[subcount + 1].length : 0); } } } } metadataList.sort(ruleMetadataComparer); for (index = 0; index < metadataList.length; index++) { sVal = replaceMetadata(sVal, metadataList[index], scheme, learningMode); } for (count = 0; count < romanizationRules.length; count++) { sVal = sVal.replace(new RegExp(romanizationRules[count].match, 'g'), (romanizationRules[count].ruleFunc === 0) ? romanizationRules[count].evaluator : replaceFunc(ruleFunctions[romanizationRules[count].ruleFunc - 1], romanizationRules[count].evaluator)); } return sVal; }"}
        NumberGenJS = {"var arabicOrdinalNumbers = " + UtilityWeb.MakeJSArray(ChData.ArabicOrdinalNumbers) + ";",
                "var arabicOrdinalExtraNumbers = " + UtilityWeb.MakeJSArray(ChData.ArabicOrdinalExtraNumbers) + ";",
                "var arabicFractionNumbers = " + UtilityWeb.MakeJSArray(ChData.ArabicFractionNumbers) + ";",
                "var arabicBaseNumbers = " + UtilityWeb.MakeJSArray(ChData.ArabicBaseNumbers) + ";",
                "var arabicBaseExtraNumbers = " + UtilityWeb.MakeJSArray(ChData.ArabicBaseExtraNumbers) + ";",
                "var arabicBaseTenNumbers = " + UtilityWeb.MakeJSArray(ChData.ArabicBaseTenNumbers) + ";",
                "var arabicBaseHundredNumbers = " + UtilityWeb.MakeJSArray(ChData.ArabicBaseHundredNumbers) + ";",
                "var arabicBaseThousandNumbers = " + UtilityWeb.MakeJSArray(ChData.ArabicBaseThousandNumbers) + ";",
                "var arabicBaseMillionNumbers = " + UtilityWeb.MakeJSArray(ChData.ArabicBaseMillionNumbers) + ";",
                "var arabicBaseBillionNumbers = " + UtilityWeb.MakeJSArray(ChData.ArabicBaseBillionNumbers) + ";",
                "var arabicBaseMilliardNumbers = " + UtilityWeb.MakeJSArray(ChData.ArabicBaseMilliardNumbers) + ";",
                "var arabicBaseTrillionNumbers = " + UtilityWeb.MakeJSArray(ChData.ArabicBaseTrillionNumbers) + ";",
                "var arabicCombiners = " + UtilityWeb.MakeJSArray(ChData.ArabicCombiners) + ";",
                "function doTransliterateNum() { $('#translitvalue').text(doTransliterate(arabicWordFromNumber($('#translitedit').val(), $('#useclassic0').prop('checked'), $('#usehundredform0').prop('checked'), $('#usemilliard0').prop('checked')), false, 1)); }",
                "function arabicWordForLessThanThousand(number, useclassic, usealefhundred) { var str = '', hundstr = ''; if (number >= 100) { hundstr = usealefhundred ? arabicBaseHundredNumbers[Math.floor(number / 100) - 1].substr(0, 2) + 'A' + arabicBaseHundredNumbers[Math.floor(number / 100) - 1].substr(2) : arabicBaseHundredNumbers[Math.floor(number / 100) - 1]; if ((number % 100) === 0) { return hundstr; } number = number % 100; } if ((number % 10) !== 0 && number !== 11 && number !== 12) { str = arabicBaseNumbers[number % 10 - 1]; } if (number >= 11 && number < 20) { if (number == 11 || number == 12) { str += arabicBaseExtraNumbers[number - 11]; } else { str = str.slice(0, -1) + 'a'; } str += ' ' + arabicBaseTenNumbers[1].slice(0, -2); } else if ((number === 0 && str === '') || number === 10 || number >= 20) { str = ((str === '') ? '' : str + ' ' + arabicCombiners[0]) + arabicBaseTenNumbers[Math.floor(number / 10)]; } return useclassic ? (((str === '') ? '' : str + ((hundstr === '') ? '' : ' ' + arabicCombiners[0])) + hundstr) : (((hundstr === '') ? '' : hundstr + ((str === '') ? '' : ' ' + arabicCombiners[0])) + str); }",
                "function arabicWordFromNumber(number, useclassic, usealefhundred, usemilliard) { var str = '', nextstr = '', curbase = 3, basenums = [1000, 1000000, 1000000000, 1000000000000], bases = [arabicBaseThousandNumbers, arabicBaseMillionNumbers, usemilliard ? arabicBaseMilliardNumbers : arabicBaseBillionNumbers, arabicBaseTrillionNumbers]; do { if (number >= basenums[curbase] && number < 2 * basenums[curbase]) { nextstr = bases[curbase][0]; } else if (number >= 2 * basenums[curbase] && number < 3 * basenums[curbase]) { nextstr = bases[curbase][1]; } else if (number >= 3 * basenums[curbase] && number < 10 * basenums[curbase]) { nextstr = arabicBaseNumbers[Math.floor(Number / basenums[curbase]) - 1].slice(0, -1) + 'u ' + bases[curbase][2]; } else if (number >= 10 * basenums[curbase] && number < 11 * basenums[curbase]) { nextstr = arabicBaseTenNumbers[1].slice(0, -1) + 'u ' + bases[curbase][2]; } else if (number >= basenums[curbase]) { nextstr = arabicWordForLessThanThousand(Math.floor(number / basenums[curbase]) % 100, useclassic, usealefhundred); if (number >= 100 * basenums[curbase] && number < (useclassic ? 200 : 101) * basenums[curbase]) { nextstr = nextstr.slice(0, -1) + 'u ' + bases[curbase][0].slice(0, -1) + 'K'; } else if (number >= 200 * basenums[curbase] && number < (useclassic ? 300 : 201) * basenums[curbase]) { nextstr = nextstr.slice(0, -2) + ' ' + bases[curbase][0].slice(0, -1) + 'K'; } else if (number >= 300 * basenums[curbase] && (useclassic || Math.floor(number / basenums[curbase]) % 100 === 0)) { nextstr = nextstr.slice(0, -1) + 'i ' + bases[curbase][0].slice(0, -1) + 'K'; } else { nextstr += ' ' + bases[curbase][0].slice(0, -1) + 'FA'; } } number = number % basenums[curbase]; curbase--; str = useclassic ? (((nextstr === '') ? '' : nextstr + ((str === '') ? '' : ' ' + arabicCombiners[0])) + str) : (((str === '') ? '' : str + ((nextstr === '') ? '' : ' ' + arabicCombiners[0])) + nextstr); nextstr = ''; } while (curbase >= 0); if (number !== 0 || str === '') { nextstr = arabicWordForLessThanThousand(number, useclassic, usealefhundred); } return useclassic ? (((nextstr === '') ? '' : nextstr + ((str === '') ? '' : ' ' + arabicCombiners[0])) + str) : (((str === '') ? '' : str + ((nextstr === '') ? '' : ' ' + arabicCombiners[0])) + nextstr); }"}
    End Sub
    Public Function DecodeTranslitScheme(Context As HttpContext) As String
        'QueryString instead of Params?
        Return Arb.DecodeTranslitScheme(Context.Request.Params("translitscheme"))
    End Function
    Public Shared Function DecodeTranslitSchemeType(Context As HttpContext) As ArabicData.TranslitScheme
        'QueryString instead of Params?
        Return Arabic.DecodeTranslitSchemeType(Context.Request.Params("translitscheme"))
    End Function
    Public Function GetTranslitSchemeMetadata(ID As String) As Array()
        Dim Output(ChData.IslamData.TranslitSchemes.Length + 2) As Array
        Output(0) = New String() {}
        Output(1) = New String() {String.Empty, String.Empty}
        Output(2) = New String() {_PortableMethods.LoadResourceString("IslamInfo_Name"), _PortableMethods.LoadResourceString("IslamInfo_Translation")}
        For Count = 0 To ChData.IslamData.TranslitSchemes.Length - 1
            Output(3 + Count) = {ChData.IslamData.TranslitSchemes(Count).Name, _PortableMethods.LoadResourceString("IslamSource_" + ChData.IslamData.TranslitSchemes(Count).Name)}
        Next
        Return RenderArrayWeb.MakeTableJSFunctions(Output, ID)
    End Function
    Public Function DisplayParticle(Category As IslamData.GrammarSet.GrammarParticle(), ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, ColSels() As String) As Array()
        Dim Count As Integer
        Dim Output(2 + Category.Length) As Array
        If ColSels Is Nothing Then ColSels = {"posspron"}
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Dim Strings As New List(Of String)
        Strings.AddRange({"arabic", "transliteration", "translation"})
        If Array.IndexOf(ColSels, "posspron") <> -1 Then
            Strings.Add(String.Empty)
        End If
        Strings.Add(String.Empty)
        Output(1) = Strings.ToArray()
        Strings.Clear()
        Strings.AddRange({"Particle", _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        If Array.IndexOf(ColSels, "posspron") <> -1 Then
            Strings.Add("Prepositional Attached Pronoun")
        End If
        Strings.Add("Grammar Feature")
        Output(2) = Strings.ToArray()
        For Count = 0 To Category.Length - 1
            Dim Objs As New List(Of Object)
            Objs.AddRange({Arb.TransliterateFromBuckwalter(Category(Count).Text), Arb.TransliterateToScheme(Arb.TransliterateFromBuckwalter(Category(Count).Text), If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(Arb.TransliterateFromBuckwalter(Category(Count).Text), Arb.GetMetarules(Arb.TransliterateFromBuckwalter(Category(Count).Text), ChData.RuleMetas("Normal")), Nothing)), _PortableMethods.LoadResourceString("IslamInfo_" + Category(Count).TranslationID)})
            If Array.IndexOf(ColSels, "posspron") <> -1 Then
                If Array.FindIndex(Utility.DefaultValue(Category(Count).Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "prep") <> -1) <> -1 Then
                    Objs.Add(DisplayTransform(Arb.TransliterateFromBuckwalter(Category(Count).Text), Arb.GetTransform("posspron"), ID, True, True, SchemeType, Scheme, Nothing))
                Else
                    Objs.Add(String.Empty)
                End If
            End If
            Objs.Add(Utility.DefaultValue(Category(Count).Grammar, String.Empty))
            Output(3 + Count) = Objs.ToArray()
        Next
        Return RenderArrayWeb.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Public Function DisplayPronoun(Category As IslamData.GrammarSet.GrammarNoun(), ID As String, Personal As Boolean, SchemeType As ArabicData.TranslitScheme, Scheme As String, ColSels() As String) As Array()
        Dim Count As Integer
        Dim Cols As String()
        Dim ColVals As String()
        If ColSels Is Nothing Then ColSels = {"p", "d", "s"}
        If Personal Then
            Cols = {"3m", "3f", "2m", "2f", "1m", "1f"}
            ColVals = {"Third Person Masculine", "Third Person Feminine", "Second Person Masculine", "Second Person Feminine", "First Person Masculine", "First Person Feminine"}
        Else
            Cols = {"m", "f"}
            ColVals = {"Masculine", "Feminine"}
        End If
        Dim Output(2 + Cols.Length) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String()))
        Output(0) = New String() {}
        Dim Strings As New List(Of String)
        If Array.IndexOf(ColSels, "p") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "d") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "s") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        Strings.Add(String.Empty)
        Output(1) = Strings.ToArray()
        Strings.Clear()
        If Array.IndexOf(ColSels, "p") <> -1 Then
            Strings.AddRange({"Plural " + _PortableMethods.LoadResourceString("IslamInfo_Arabic"), "Plural " + _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), "Plural " + _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "d") <> -1 Then
            Strings.AddRange({"Dual " + _PortableMethods.LoadResourceString("IslamInfo_Arabic"), "Dual " + _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), "Dual " + _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "s") <> -1 Then
            Strings.AddRange({"Singular " + _PortableMethods.LoadResourceString("IslamInfo_Arabic"), "Singular " + _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), "Singular " + _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        Strings.Add("Person and Gender")
        Output(2) = Strings.ToArray()
        Strings.Clear()
        For Count = 0 To Category.Length - 1
            Dim Translat As String = _PortableMethods.LoadResourceString("IslamInfo_" + Category(Count).TranslationID)
            For Each Sets As String In Category(Count).Grammar.Split(","c)
                For Each Str As String In Sets.Split("|"c)
                    If System.Text.RegularExpressions.Regex.Match(Str, "^(?:" + ArabicData.MakeRegMultiEx(Cols) + ")[pds]$").Success Then
                        Dim Key As String = Str.Chars(0)
                        If Personal Then '"123".Contains(Str.Chars(0))
                            Key += Str.Chars(1)
                        End If
                        If Not Build.ContainsKey(Key) Then
                            Build.Add(Key, New Generic.Dictionary(Of String, String()))
                        End If
                        If Build.Item(Key).ContainsKey(Str.Chars(If(Personal, 2, 1))) Then
                            Build.Item(Key).Item(Str.Chars(If(Personal, 2, 1)))(0) += " " + Arb.TransliterateFromBuckwalter(Category(Count).Text)
                        Else
                            Build.Item(Key).Add(Str.Chars(If(Personal, 2, 1)), {Arb.TransliterateFromBuckwalter(Category(Count).Text), Translat})
                        End If
                    End If
                Next
            Next
        Next
        For Index = 0 To Cols.Length - 1
            If Build.ContainsKey(Cols(Index)) Then
                Dim Strs(3 * ColSels.Length) As String
                For SubIndex = 0 To ColSels.Length - 1
                    If Not Build(Cols(Index)).ContainsKey(ColSels(SubIndex)) Then Build(Cols(Index)).Add(ColSels(SubIndex), {String.Empty, String.Empty})
                    Strs(3 * SubIndex) = Build(Cols(Index))(ColSels(SubIndex))(0)
                    Strs(3 * SubIndex + 1) = Arb.TransliterateToScheme(Build(Cols(Index))(ColSels(SubIndex))(0), If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(Build(Cols(Index))(ColSels(SubIndex))(0), Arb.GetMetarules(Build(Cols(Index))(ColSels(SubIndex))(0), ChData.RuleMetas("Normal")), Nothing))
                    Strs(3 * SubIndex + 2) = Build(Cols(Index))(ColSels(SubIndex))(1)
                Next
                Strs(3 * ColSels.Length) = ColVals(Index)
                Output(3 + Index) = Strs
            End If
        Next
        Return RenderArrayWeb.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Public Function DisplayTransform(Text As String, Category As IslamData.GrammarSet.GrammarTransform(), ID As String, Personal As Boolean, Noun As Boolean, SchemeType As ArabicData.TranslitScheme, Scheme As String, ColSels As String()) As Array()
        Dim Count As Integer
        Dim Cols As String()
        Dim ColVals As String()
        If ColSels Is Nothing Then ColSels = {"p", "d", "s"}
        If Personal Then
            Cols = {"3m", "3f", "2m", "2f", "1m", "1f"}
            ColVals = {"Third Person Masculine", "Third Person Feminine", "Second Person Masculine", "Second Person Feminine", "First Person Masculine", "First Person Feminine"}
        Else
            Cols = {"m", "f"}
            ColVals = {"Masculine", "Feminine"}
        End If
        Dim Output(2 + Cols.Length) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String()))
        Output(0) = New String() {}
        Dim Strings As New List(Of String)
        If Array.IndexOf(ColSels, "p") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "d") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "s") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        Strings.Add(String.Empty)
        Output(1) = Strings.ToArray()
        Strings.Clear()
        If Array.IndexOf(ColSels, "p") <> -1 Then
            Strings.AddRange({"Plural " + _PortableMethods.LoadResourceString("IslamInfo_Arabic"), "Plural " + _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), "Plural " + _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "d") <> -1 Then
            Strings.AddRange({"Dual " + _PortableMethods.LoadResourceString("IslamInfo_Arabic"), "Dual " + _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), "Dual " + _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "s") <> -1 Then
            Strings.AddRange({"Singular " + _PortableMethods.LoadResourceString("IslamInfo_Arabic"), "Singular " + _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), "Singular " + _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        Strings.Add("Person and Gender")
        Output(2) = Strings.ToArray()
        Strings.Clear()
        For Count = 0 To Category.Length - 1
            Dim Translat As String = _PortableMethods.LoadResourceString("IslamInfo_" + Category(Count).TranslationID)
            For Each Sets As String In Category(Count).Grammar.Split(","c)
                For Each Str As String In Sets.Split("|"c)
                    If Array.FindIndex(Category(Count).Grammar.Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), If(Noun, "noun", "verb")) <> -1) <> -1 And System.Text.RegularExpressions.Regex.Match(Str, "^(?:" + ArabicData.MakeRegMultiEx(Cols) + ")[pds]$").Success Then
                        Dim Key As String = Str.Chars(0)
                        If Personal Then '"123".Contains(Str.Chars(0))
                            Key += Str.Chars(1)
                        End If
                        If Not Build.ContainsKey(Key) Then
                            Build.Add(Key, New Generic.Dictionary(Of String, String()))
                        End If
                        If Build.Item(Key).ContainsKey(Str.Chars(If(Personal, 2, 1))) Then
                            Build.Item(Key).Item(Str.Chars(If(Personal, 2, 1)))(0) = If(Text = String.Empty, Build.Item(Key).Item(Str.Chars(If(Personal, 2, 1)))(0) + " " + ChData.TranslateRegEx(Category(Count).Text, False), " " + Arb.ApplyTransform({Category(Count)}, Build.Item(Key).Item(Str.Chars(If(Personal, 2, 1)))(0)))
                            Build.Item(Key).Item(Str.Chars(If(Personal, 2, 1)))(1) = Translat
                        Else
                            Build.Item(Key).Add(Str.Chars(If(Personal, 2, 1)), {If(Text = String.Empty, ChData.TranslateRegEx(Category(Count).Text, False), Arb.ApplyTransform({Category(Count)}, Text)), Translat})
                        End If
                    End If
                Next
            Next
        Next
        For Index = 0 To Cols.Length - 1
            If Build.ContainsKey(Cols(Index)) Then
                Dim Strs(3 * ColSels.Length) As String
                For SubIndex = 0 To ColSels.Length - 1
                    If Not Build(Cols(Index)).ContainsKey(ColSels(SubIndex)) Then Build(Cols(Index)).Add(ColSels(SubIndex), {String.Empty, String.Empty})
                    Strs(3 * SubIndex) = Build(Cols(Index))(ColSels(SubIndex))(0)
                    Strs(3 * SubIndex + 1) = Arb.TransliterateToScheme(Build(Cols(Index))(ColSels(SubIndex))(0), If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(Build(Cols(Index))(ColSels(SubIndex))(0), Arb.GetMetarules(Build(Cols(Index))(ColSels(SubIndex))(0), ChData.RuleMetas("Normal")), Nothing))
                    Strs(3 * SubIndex + 2) = Build(Cols(Index))(ColSels(SubIndex))(1)
                Next
                Strs(3 * ColSels.Length) = ColVals(Index)
                Output(3 + Index) = Strs
            End If
        Next
        Return RenderArrayWeb.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Public Function DisplayWord(Category As IslamData.GrammarSet.GrammarWord(), ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Dim Count As Integer
        Dim Output(2 + Category.Length) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", "transliteration", "translation", String.Empty}
        Output(2) = New String() {_PortableMethods.LoadResourceString("IslamInfo_Arabic"), _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), _PortableMethods.LoadResourceString("IslamInfo_Translation"), _PortableMethods.LoadResourceString("IslamInfo_Grammar")}
        For Count = 0 To Category.Length - 1
            Output(3 + Count) = New String() {Arb.TransliterateFromBuckwalter(Category(Count).Text), Arb.TransliterateToScheme(Arb.TransliterateFromBuckwalter(Category(Count).Text), If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(Arb.TransliterateFromBuckwalter(Category(Count).Text), Arb.GetMetarules(Arb.TransliterateFromBuckwalter(Category(Count).Text), ChData.RuleMetas("Normal")), Nothing)), _PortableMethods.LoadResourceString("IslamInfo_" + Category(Count).TranslationID), If(Category(Count).Grammar Is Nothing, String.Empty, Category(Count).Grammar)}
        Next
        Return RenderArrayWeb.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Public Async Function DisplayDict(ByVal Item As PageLoader.TextItem) As Task(Of Array())
        Return Await Arb.DisplayDict()
    End Function
    Public Function DisplayCombo(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Return Arb.DisplayCombo(DecodeTranslitSchemeType(Context), DecodeTranslitScheme(Context))
    End Function
    Public Function DisplayAll(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Return Arb.DisplayAll(DecodeTranslitSchemeType(Context), DecodeTranslitScheme(Context))
    End Function
    Public Function DisplayTranslitSchemes(ByVal Item As PageLoader.TextItem) As Array()
        Return Arb.DisplayTranslitSchemes()
    End Function
    Public Function DisplayProximals(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayPronoun(Arb.GetCatNoun("proxdemo"), Item.Name, False, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayDistals(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayPronoun(Arb.GetCatNoun("distdemo"), Item.Name, False, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayRelatives(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayPronoun(Arb.GetCatNoun("relpro"), Item.Name, False, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayPersonals(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayPronoun(Arb.GetCatNoun("perspro"), Item.Name, True, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayDeterminerPersonals(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayTransform(String.Empty, Arb.GetTransform("posspron"), Item.Name, True, True, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayPastVerbsFamilyI(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayTransform(String.Empty, Arb.GetTransform("pastverbi"), Item.Name, True, False, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayPresentVerbsFamilyI(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayTransform(String.Empty, Arb.GetTransform("presverbi"), Item.Name, True, False, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayCommandVerbsFamilyI(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayTransform(String.Empty, Arb.GetTransform("commverbi"), Item.Name, False, False, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayResponseParticles(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayParticle(Arb.GetParticles("resp"), Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayInterogativeParticles(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayParticle(Arb.GetParticles("intg"), Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayLocationParticles(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayParticle(Arb.GetParticles("loc"), Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayTimeParticles(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayParticle(Arb.GetParticles("time"), Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayPrepositions(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayParticle(Arb.GetParticles("prep"), Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayParticles(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return DisplayParticle(Arb.GetParticles("particle"), Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Function NounDisplay(Category As IslamData.GrammarSet.GrammarNoun(), ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, ColSels As String()) As Array()
        Dim Count As Integer
        Dim Output(2 + Category.Length) As Array
        Output(0) = New String() {}
        If ColSels Is Nothing Then ColSels = {"deft", "fem", "reladj", "grammar"}
        Dim Strings As New List(Of String)
        Strings.Add(String.Empty)
        If Array.IndexOf(ColSels, "deft") <> -1 Then
            Strings.Add(String.Empty)
        End If
        If Array.IndexOf(ColSels, "fem") <> -1 Then
            Strings.Add(String.Empty)
        End If
        If Array.IndexOf(ColSels, "reladj") <> -1 Then
            Strings.Add(String.Empty)
        End If
        If Array.IndexOf(ColSels, "grammar") <> -1 Then
            Strings.Add(String.Empty)
        End If
        Output(1) = Strings.ToArray()
        Strings.Clear()
        Strings.Add("Base Noun")
        If Array.IndexOf(ColSels, "deft") <> -1 Then
            Strings.Add("Definite")
        End If
        If Array.IndexOf(ColSels, "fem") <> -1 Then
            Strings.Add("Feminine")
        End If
        If Array.IndexOf(ColSels, "reladj") <> -1 Then
            Strings.Add("Relational Adjective")
        End If
        If Array.IndexOf(ColSels, "grammar") <> -1 Then
            Strings.Add("Grammar")
        End If
        Output(2) = Strings.ToArray()
        Strings.Clear()
        For Count = 0 To Category.Length - 1
            Dim Tables As New List(Of Object)
            Tables.Add(DeclineNoun(Category(Count), ID, SchemeType, Scheme, ColSels))
            If Array.IndexOf(ColSels, "deft") <> -1 Then
                Dim Text As String = Arb.ApplyTransform(Arb.GetTransform("deft"), Arb.ApplyTransform(Arb.GetTransform("strip"), Arb.TransliterateFromBuckwalter(Category(Count).Text)))
                Tables.Add(DeclineNoun(New IslamData.GrammarSet.GrammarNoun With {.Text = Text, .Grammar = "flex,def," + If(Array.FindIndex(Utility.DefaultValue(Category(Count).Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, "fs", "ms") + If(Array.FindIndex(Utility.DefaultValue(Category(Count).Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "reladj") <> -1) <> -1, ",reladj", String.Empty), .TranslationID = Category(Count).TranslationID}, ID, SchemeType, Scheme, ColSels))
            End If
            If Array.IndexOf(ColSels, "fem") <> -1 Then
                If Array.FindIndex(Utility.DefaultValue(Category(Count).Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) = -1 Then
                    Dim Text As String = Arb.ApplyTransform(Arb.GetTransform("fem"), Arb.ApplyTransform(Arb.GetTransform("strip"), Arb.TransliterateFromBuckwalter(Category(Count).Text)))
                    Tables.Add(NounDisplay({New IslamData.GrammarSet.GrammarNoun With {.Text = Text, .Grammar = "flex,indef,fs" + If(Array.FindIndex(Utility.DefaultValue(Category(Count).Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "reladj") <> -1) <> -1, ",reladj", String.Empty), .TranslationID = Category(Count).TranslationID}}, ID, SchemeType, Scheme, ColSels))
                Else
                    Tables.Add(String.Empty)
                End If
            End If
            If Array.IndexOf(ColSels, "reladj") <> -1 Then
                If Array.FindIndex(Utility.DefaultValue(Category(Count).Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "adj") <> -1 Or Array.IndexOf(S.Split("|"c), "reladj") <> -1 Or Array.IndexOf(S.Split("|"c), "fs") <> -1) = -1 Then
                    Dim Text As String = Arb.ApplyTransform(Arb.GetTransform("reladj"), Arb.ApplyTransform(Arb.GetTransform("strip"), Arb.TransliterateFromBuckwalter(Category(Count).Text)))
                    Tables.Add(NounDisplay({New IslamData.GrammarSet.GrammarNoun With {.Text = Text, .Grammar = "flex,reladj,indef,ms", .TranslationID = Category(Count).TranslationID}}, ID, SchemeType, Scheme, ColSels))
                Else
                    Tables.Add(String.Empty)
                End If
            End If
            If Array.IndexOf(ColSels, "grammar") <> -1 Then
                Tables.Add(Category(Count).Grammar)
            End If
            Output(3 + Count) = Tables.ToArray()
            Tables.Clear()
        Next
        Return RenderArrayWeb.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Public Function DisplayNouns(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return NounDisplay(ChData.IslamData.Grammar.Nouns, Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Function DeclineNoun(Category As IslamData.GrammarSet.GrammarNoun, ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, ColSels As String()) As Array()
        Dim Sels As String() = {"nom", "acc", "pos"}
        Dim SelTexts As String() = {"Nominative", "Accusative", "Possessive"}
        Dim Output(2 + Sels.Length) As Array
        If ColSels Is Nothing Then ColSels = {"s", "d", "p", "posspron"}
        Dim HasPoss As Boolean = Not Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1 And Array.IndexOf(ColSels, "posspron") <> -1
        Output(0) = New String() {}
        Dim Strings As New List(Of String)
        If Array.IndexOf(ColSels, "s") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
            If HasPoss Then Strings.Add(String.Empty)
        End If
        If Array.IndexOf(ColSels, "d") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
            If HasPoss Then Strings.Add(String.Empty)
        End If
        If Array.IndexOf(ColSels, "p") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
            If HasPoss Then Strings.Add(String.Empty)
        End If
        Strings.Add(String.Empty)
        Output(1) = Strings.ToArray()
        Strings.Clear()
        If Array.IndexOf(ColSels, "s") <> -1 Then
            Strings.AddRange({"Singular " + _PortableMethods.LoadResourceString("IslamInfo_Arabic"), "Singular " + _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), "Singular " + _PortableMethods.LoadResourceString("IslamInfo_Translation")})
            If HasPoss Then Strings.Add("Attached Possessive Pronoun")
        End If
        If Array.IndexOf(ColSels, "d") <> -1 Then
            Strings.AddRange({"Dual " + _PortableMethods.LoadResourceString("IslamInfo_Arabic"), "Dual " + _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), "Dual " + _PortableMethods.LoadResourceString("IslamInfo_Translation")})
            If HasPoss Then Strings.Add("Attached Possessive Pronoun")
        End If
        If Array.IndexOf(ColSels, "p") <> -1 Then
            Strings.AddRange({"Plural " + _PortableMethods.LoadResourceString("IslamInfo_Arabic"), "Plural " + _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), "Plural " + _PortableMethods.LoadResourceString("IslamInfo_Translation")})
            If HasPoss Then Strings.Add("Attached Possessive Pronoun")
        End If
        Strings.Add("Case")
        Output(2) = Strings.ToArray()
        Strings.Clear()
        For Count = 0 To Sels.Length - 1
            Dim Objs As New List(Of Object)
            'Nisbah has a whole slow of suffix possibilities from like -ese or -ism or -ist -ar
            If Array.IndexOf(ColSels, "s") <> -1 Then
                Dim Text As String = Arb.ApplyTransform(Arb.GetTransformMatch({"flex", Sels(Count), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1, "def", "indef"), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, "fs", "ms")}), Arb.ApplyTransform(Arb.GetTransform("strip"), Arb.TransliterateFromBuckwalter(Category.Text)))
                Objs.AddRange({Text, Arb.TransliterateToScheme(Text, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(Text, Arb.GetMetarules(Text, ChData.RuleMetas("Normal")), Nothing)), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "reladj") <> -1) <> -1, "Relating to ", String.Empty) + If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1, "The ", String.Empty) + _PortableMethods.LoadResourceString("IslamInfo_" + Category.TranslationID) + If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, " Feminine", " Masculine")})
                If HasPoss Then
                    Objs.Add(DisplayTransform(Arb.ApplyTransform(Arb.GetTransform("constpos"), Text), Arb.GetTransform("posspron"), ID, True, True, SchemeType, Scheme, Array.FindAll(ColSels, Function(S As String) Array.IndexOf({"p", "d", "s"}, S) <> -1)))
                End If
            End If
            If Array.IndexOf(ColSels, "d") <> -1 Then
                Dim Text As String = Arb.ApplyTransform(Arb.GetTransformMatch({"flex", Sels(Count), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1, "def", "indef"), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, "fd", "md")}), Arb.ApplyTransform(Arb.GetTransform("strip"), Arb.TransliterateFromBuckwalter(Category.Text)))
                Objs.AddRange({Text, Arb.TransliterateToScheme(Text, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(Text, Arb.GetMetarules(Text, ChData.RuleMetas("Normal")), Nothing)), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "reladj") <> -1) <> -1, "Relating to ", String.Empty) + If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1, "The ", String.Empty) + "Two " + _PortableMethods.LoadResourceString("IslamInfo_" + Category.TranslationID) + "s" + If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, " Feminine", " Masculine")})
                If HasPoss Then
                    Objs.Add(DisplayTransform(Arb.ApplyTransform(Arb.GetTransform("constpos"), Text), Arb.GetTransform("posspron"), ID, True, True, SchemeType, Scheme, Array.FindAll(ColSels, Function(S As String) Array.IndexOf({"p", "d", "s"}, S) <> -1)))
                End If
            End If
            If Array.IndexOf(ColSels, "p") <> -1 Then
                Dim Text As String = Arb.ApplyTransform(Arb.GetTransformMatch({"flex", Sels(Count), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1, "def", "indef"), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, "fp", "mp")}), Arb.ApplyTransform(Arb.GetTransform("strip"), Arb.TransliterateFromBuckwalter(Category.Text)))
                Objs.AddRange({Text, Arb.TransliterateToScheme(Text, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(Text, Arb.GetMetarules(Text, ChData.RuleMetas("Normal")), Nothing)), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "reladj") <> -1) <> -1, "Relating to ", String.Empty) + If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1, "The ", String.Empty) + _PortableMethods.LoadResourceString("IslamInfo_" + Category.TranslationID) + "s" + If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, " Feminine", " Masculine")})
                If HasPoss Then
                    Objs.Add(DisplayTransform(Arb.ApplyTransform(Arb.GetTransform("constpos"), Text), Arb.GetTransform("posspron"), ID, True, True, SchemeType, Scheme, Array.FindAll(ColSels, Function(S As String) Array.IndexOf({"p", "d", "s"}, S) <> -1)))
                End If
            End If
            Objs.Add(SelTexts(Count))
            Output(3 + Count) = Objs.ToArray()
            Objs.Clear()
        Next
        Return RenderArrayWeb.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Public Function VerbDisplay(Category As IslamData.GrammarSet.GrammarVerb(), ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, ColSels As String()) As Array()
        Dim Count As Integer
        Dim Output(2 + Category.Length) As Array
        If ColSels Is Nothing Then ColSels = {"past", "pres", "comm", "forbid", "pasvpast", "pasvpres", "doernoun", "pasvnoun", "part"}
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Dim Strings As New List(Of String)
        If Array.IndexOf(ColSels, "past") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "pres") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "comm") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "forbid") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "pasvpast") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "pasvpres") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "doernoun") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "pasvnoun") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "part") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        Output(1) = Strings.ToArray()
        Strings.Clear()
        If Array.IndexOf(ColSels, "past") <> -1 Then
            Strings.AddRange({"Past Root", _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "pres") <> -1 Then
            Strings.AddRange({"Present Root", _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "comm") <> -1 Then
            Strings.AddRange({"Command Root", _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "forbid") <> -1 Then
            Strings.AddRange({"Forbidding Root", _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "pasvpast") <> -1 Then
            Strings.AddRange({"Passive Past Root", _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "pasvpres") <> -1 Then
            Strings.AddRange({"Passive Present Root", _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "doernoun") <> -1 Then
            Strings.AddRange({"Verbal Doer", _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "pasvnoun") <> -1 Then
            Strings.AddRange({"Passive Noun", _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "part") <> -1 Then
            Strings.AddRange({"Particles", _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), _PortableMethods.LoadResourceString("IslamInfo_Translation")})
        End If
        Output(2) = Strings.ToArray()
        Strings.Clear()
        For Count = 0 To Category.Length - 1
            Dim Grammar As String
            Dim Text As String
            Dim Present As String
            Dim Command As String
            Dim Forbidding As String
            Dim PassivePast As String
            Dim PassivePresent As String
            Dim VerbalDoer As String
            Dim PassiveNoun As String
            If (Not Category(Count).Grammar Is Nothing AndAlso Category(Count).Grammar.StartsWith("form=")) Then
                Text = Arb.TransliterateFromBuckwalter(Category(Count).Grammar.Substring(5).Split(","c)(0).Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)))
                Present = Arb.TransliterateFromBuckwalter(Category(Count).Grammar.Substring(5).Split(","c)(1).Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)))
                Command = Arb.TransliterateFromBuckwalter(Arb.GetTransform("VerbTypeICommandYouMasculinePattern")(0).Text.Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)).Replace("\1", Category(Count).Grammar.Substring(5).Split(","c)(1).Chars(5)))
                Dim Multi As String() = Arb.GetTransform("VerbTypeIForbiddingYouMasculinePattern")(0).Text.Split(" "c)
                Forbidding = Arb.TransliterateFromBuckwalter(Multi(0) + " " + Multi(1).Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)).Replace("\1", Category(Count).Grammar.Substring(5).Split(","c)(1).Chars(5)))
                PassivePast = Arb.TransliterateFromBuckwalter(Arb.GetTransform("VerbTypeIPassivePastHePattern")(0).Text.Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)))
                PassivePresent = Arb.TransliterateFromBuckwalter(Arb.GetTransform("VerbTypeIPassivePresentHeMasculinePattern")(0).Text.Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)))
                VerbalDoer = Arb.TransliterateFromBuckwalter(Arb.GetTransform("VerbTypeIVerbalDoerPattern")(0).Text.Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)))
                PassiveNoun = Arb.TransliterateFromBuckwalter(Arb.GetTransform("VerbTypeIPassiveNounPattern")(0).Text.Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)))
                Grammar = String.Empty
            Else
                Text = Arb.TransliterateFromBuckwalter(Category(Count).Text)
                Present = String.Empty
                Command = String.Empty
                Forbidding = String.Empty
                PassivePast = String.Empty
                PassivePresent = String.Empty
                VerbalDoer = String.Empty
                PassiveNoun = String.Empty
                Grammar = Utility.DefaultValue(Category(Count).Grammar, String.Empty)
            End If
            If Array.IndexOf(ColSels, "past") <> -1 Then
                Strings.AddRange({Text, Arb.TransliterateToScheme(Text, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(Text, Arb.GetMetarules(Text, ChData.RuleMetas("Normal")), Nothing)), _PortableMethods.LoadResourceString("IslamInfo_" + Category(Count).TranslationID)})
            End If
            If Array.IndexOf(ColSels, "pres") <> -1 Then
                Strings.AddRange({Present, Arb.TransliterateToScheme(Present, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(Present, Arb.GetMetarules(Present, ChData.RuleMetas("Normal")), Nothing)), String.Empty})
            End If
            If Array.IndexOf(ColSels, "comm") <> -1 Then
                Strings.AddRange({Command, Arb.TransliterateToScheme(Command, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(Command, Arb.GetMetarules(Command, ChData.RuleMetas("Normal")), Nothing)), String.Empty})
            End If
            If Array.IndexOf(ColSels, "forbid") <> -1 Then
                Strings.AddRange({Forbidding, Arb.TransliterateToScheme(Forbidding, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(Forbidding, Arb.GetMetarules(Forbidding, ChData.RuleMetas("Normal")), Nothing)), String.Empty})
            End If
            If Array.IndexOf(ColSels, "pasvpast") <> -1 Then
                Strings.AddRange({PassivePast, Arb.TransliterateToScheme(PassivePast, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(PassivePast, Arb.GetMetarules(PassivePast, ChData.RuleMetas("Normal")), Nothing)), String.Empty})
            End If
            If Array.IndexOf(ColSels, "pasvpres") <> -1 Then
                Strings.AddRange({PassivePresent, Arb.TransliterateToScheme(PassivePresent, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(PassivePresent, Arb.GetMetarules(PassivePresent, ChData.RuleMetas("Normal")), Nothing)), String.Empty})
            End If
            If Array.IndexOf(ColSels, "doernoun") <> -1 Then
                Strings.AddRange({VerbalDoer, Arb.TransliterateToScheme(VerbalDoer, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(VerbalDoer, Arb.GetMetarules(VerbalDoer, ChData.RuleMetas("Normal")), Nothing)), String.Empty})
            End If
            If Array.IndexOf(ColSels, "pasvnoun") <> -1 Then
                Strings.AddRange({PassiveNoun, Arb.TransliterateToScheme(PassiveNoun, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(PassiveNoun, Arb.GetMetarules(PassiveNoun, ChData.RuleMetas("Normal")), Nothing)), String.Empty})
            End If
            If Array.IndexOf(ColSels, "part") <> -1 Then
                Strings.AddRange({Grammar, Arb.TransliterateToScheme(Grammar, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(Grammar, Arb.GetMetarules(Grammar, ChData.RuleMetas("Normal")), Nothing)), String.Empty})
            End If
            Output(3 + Count) = Strings.ToArray()
            Strings.Clear()
        Next
        Return RenderArrayWeb.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Public Function DisplayVerbs(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Return VerbDisplay(ChData.IslamData.Grammar.Verbs, Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayFirstFormVerbs(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Object()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Dim Output As New List(Of Object)
        Output.Add(New String() {})
        Output.Add(New String() {"arabic", "arabic", "arabic", "arabic", "arabic", "arabic", "arabic", String.Empty, String.Empty, String.Empty})
        Output.Add(New String() {"Root", "Past Middle Vowel", "Present Middle Vowel", "Other Present Middle Vowel", "Past", "Present", "Other Present", "Past Verses", "Present Verses", "Other Present Verses"})
        Dim TR As New TanzilReader(_PortableMethods, Arb, ArbData, ChData)
        Output.AddRange(ChData.GetMorphologicalDataByVerbScale(TR))
        Return Output.ToArray()
    End Function
    Public Function LettersByChapter(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Object()
        Dim SchemeType As ArabicData.TranslitScheme = DecodeTranslitSchemeType(Context)
        Dim Scheme As String = DecodeTranslitScheme(Context)
        Dim Output As New List(Of String())
        Output.Add(New String() {})
        Dim TR As New TanzilReader(_PortableMethods, Arb, ArbData, ChData)
        ChData.BuildQuranLetterIndex(TR)
        Dim Verses As Collections.Generic.List(Of String())
        Verses = TR.GetQuranText(ChData.XMLDocMain, -1, -1, -1, -1)
        'AlefWasla, Alef - silent, pronounced, AlefMaksura, SuperscriptAlef - always, can be written with alef, Waw, Yah, SmallWaw, SmallYah
        Output.Add(New String() {"transliteration",
                   "translation", "translation", "translation",
                   "translation", "translation", "translation",
                   "translation", "translation", "translation",
                   "translation", "translation", "translation",
                   "translation", "translation", "translation",
                   "translation", "translation", "translation",
                   "translation", "translation", "translation",
                   "translation", "translation", "translation",
                   "translation", "translation", "translation",
                   "translation", "translation", "translation",
                   "translation", "translation", "translation", "translation",
                   "translation", "translation", "translation",
                   "translation", "translation", "translation",
                   "translation", "translation", "translation"})
        Output.Add(New String() {_PortableMethods.LoadResourceString("IslamInfo_Chapter"),
                   ArbData.GetUnicodeName(ArabicData.ArabicLetterAlefWasla), ArbData.GetUnicodeName(ArabicData.ArabicLetterAlef), ArbData.GetUnicodeName(ArabicData.ArabicLetterSuperscriptAlef),
                   ArbData.GetUnicodeName(ArabicData.ArabicLetterAlefMaksura), ArbData.GetUnicodeName(ArabicData.ArabicLetterWaw), ArbData.GetUnicodeName(ArabicData.ArabicSmallWaw),
                   ArbData.GetUnicodeName(ArabicData.ArabicLetterYeh), ArbData.GetUnicodeName(ArabicData.ArabicSmallYeh), ArbData.GetUnicodeName(ArabicData.ArabicLetterBeh),
                   ArbData.GetUnicodeName(ArabicData.ArabicLetterTeh), ArbData.GetUnicodeName(ArabicData.ArabicLetterTheh), ArbData.GetUnicodeName(ArabicData.ArabicLetterJeem),
                   ArbData.GetUnicodeName(ArabicData.ArabicLetterHah), ArbData.GetUnicodeName(ArabicData.ArabicLetterKhah), ArbData.GetUnicodeName(ArabicData.ArabicLetterDal),
                   ArbData.GetUnicodeName(ArabicData.ArabicLetterThal), ArbData.GetUnicodeName(ArabicData.ArabicLetterReh), ArbData.GetUnicodeName(ArabicData.ArabicLetterZain),
                   ArbData.GetUnicodeName(ArabicData.ArabicLetterSeen), ArbData.GetUnicodeName(ArabicData.ArabicLetterSheen), ArbData.GetUnicodeName(ArabicData.ArabicLetterSad),
                   ArbData.GetUnicodeName(ArabicData.ArabicLetterDad), ArbData.GetUnicodeName(ArabicData.ArabicLetterTah), ArbData.GetUnicodeName(ArabicData.ArabicLetterZah),
                   ArbData.GetUnicodeName(ArabicData.ArabicLetterAin), ArbData.GetUnicodeName(ArabicData.ArabicLetterGhain), ArbData.GetUnicodeName(ArabicData.ArabicLetterFeh),
                   ArbData.GetUnicodeName(ArabicData.ArabicLetterQaf), ArbData.GetUnicodeName(ArabicData.ArabicLetterKaf), ArbData.GetUnicodeName(ArabicData.ArabicLetterLam),
                   ArbData.GetUnicodeName(ArabicData.ArabicLetterMeem), ArbData.GetUnicodeName(ArabicData.ArabicLetterNoon), ArbData.GetUnicodeName(ArabicData.ArabicLetterHeh), ArbData.GetUnicodeName(ArabicData.ArabicLetterTehMarbuta),
                   ArbData.GetUnicodeName(ArabicData.ArabicLetterHamza), ArbData.GetUnicodeName(ArabicData.ArabicHamzaAbove), ArbData.GetUnicodeName(ArabicData.ArabicLetterAlefWithHamzaAbove),
                   ArbData.GetUnicodeName(ArabicData.ArabicLetterAlefWithHamzaBelow), ArbData.GetUnicodeName(ArabicData.ArabicLetterWawWithHamzaAbove), ArbData.GetUnicodeName(ArabicData.ArabicLetterYehWithHamzaAbove),
                   "Total Hamzas", "Alef Maksura -> Alef", "AlefMaksura -> Yaa"})
        Dim Count As Integer
        For Count = 1 To TR.GetChapterCount()
            Output.Add(New String() {TR.GetChapterName(TR.GetChapterByIndex(Count), SchemeType, Scheme),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterAlefWasla), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterAlef), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterSuperscriptAlef), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterAlefMaksura), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterWaw), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicSmallWaw), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterYeh), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicSmallYeh), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterBeh), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterTeh), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterTheh), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterJeem), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterHah), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterKhah), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterDal), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterThal), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterReh), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterZain), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterSeen), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterSheen), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterSad), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterDad), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterTah), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterZah), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterAin), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterGhain), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterFeh), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterQaf), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterKaf), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterLam), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterMeem), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterNoon), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterHeh), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterTehMarbuta), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterHamza), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicHamzaAbove), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterAlefWithHamzaAbove), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterAlefWithHamzaBelow), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterWawWithHamzaAbove), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterYehWithHamzaAbove), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterHamza), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count()) + Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicHamzaAbove), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count()) + Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterAlefWithHamzaAbove), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count()) + Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterAlefWithHamzaBelow), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count()) + Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterWawWithHamzaAbove), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count()) + Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterYehWithHamzaAbove), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterAlefMaksura), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1 AndAlso Verses(Ints(0))(Ints(1))(Ints(2)) = ArabicData.ArabicFatha).Count())),
                       CStr(Linq.Enumerable.Sum(ChData.LetterDictionary(ArabicData.ArabicLetterAlefMaksura), Function(KV) Linq.Enumerable.Where(KV.Value, Function(Ints) Ints(0) = Count - 1 AndAlso Verses(Ints(0))(Ints(1))(Ints(2)) = ArabicData.ArabicKasra).Count()))})
        Next
        Return Output.ToArray()
    End Function
    Public Function GetChangeTransliterationJS() As String()
        Dim GetJS As New List(Of String) From {"javascript: changeTransliteration();", String.Empty, UtilityWeb.GetLookupStyleSheetJS(), GetArabicSymbolJSArray(), GetTranslitSchemeJSArray(),
        "function processTransliteration(list) { var k, child, iSubCount, text; $('span.transliteration').each(function() { $(this).css('display', $('#translitscheme').val() === '0' ? 'none' : 'block'); }); for (k in list) { text = ''; if (list.hasOwnProperty(k) && list[k]['linkchild']) { for (child in list[k]['children']) { if (list[k]['children'].hasOwnProperty(child)) { processTransliteration(list[k]['children'][child]['children']); for (iSubCount = 0; iSubCount < list[k]['children'][child]['arabic'].length; iSubCount++) { if ($('#translitscheme').val() !== '0' && $('#translitscheme').val() !== '1'  && parseInt($('#translitscheme').val(), 10) % 2 !== 1 && list[k]['children'][child]['arabic'][iSubCount] !== '' && list[k]['children'][child]['translit'][iSubCount] !== '') { if (text !== '') text += ' '; text += $('#' + list[k]['children'][child]['arabic'][iSubCount]).text(); } else { if (list[k]['children'][child]['translit'][iSubCount] !== '') $('#' + list[k]['children'][child]['translit'][iSubCount]).text(($('#translitscheme').val() === '0' || list[k]['children'][child]['arabic'][iSubCount] === '') ? '' : doTransliterate($('#' + list[k]['children'][child]['arabic'][iSubCount]).text(), true, parseInt($('#translitscheme').val(), 10))); } } } } if ($('#translitscheme').val() !== '0' && $('#translitscheme').val() !== '1' && parseInt($('#translitscheme').val(), 10) % 2 !== 1) { text = transliterateWithRules(text, Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2, null, false).split(' '); for (child in list[k]['children']) { if (list[k]['children'].hasOwnProperty(child)) { for (iSubCount = 0; iSubCount < list[k]['children'][child]['translit'].length; iSubCount++) { if (list[k]['children'][child]['arabic'][iSubCount] !== '' && list[k]['children'][child]['translit'][iSubCount] !== '') $('#' + list[k]['children'][child]['translit'][iSubCount]).text(text.shift()); } } } } } else { processTransliteration(list[k]['children']); } for (iSubCount = 0; iSubCount < list[k]['arabic'].length; iSubCount++) { if (list[k]['translit'][iSubCount] !== '') $('#' + list[k]['translit'][iSubCount]).text(($('#translitscheme').val() === '0' || list[k]['arabic'][iSubCount] === '') ? '' : (($('#translitscheme').val() !== '0' && $('#translitscheme').val() !== '1' && parseInt($('#translitscheme').val(), 10) % 2 !== 1) ? transliterateWithRules($('#' + list[k]['arabic'][iSubCount]).text(), parseInt($('#translitscheme').val(), 10) >= 2 ? Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2 : parseInt($('#translitscheme').val(), 10), null, false) : doTransliterate($('#' + list[k]['arabic'][iSubCount]).text(), true, parseInt($('#translitscheme').val(), 10)))); } } }",
        "function changeTransliteration() { changeChapterTranslit(); var i; for (i = 0; i < renderList.length; i++) { processTransliteration(renderList[i]); } }",
        "function changeChapterTranslit() { var i; for (i = 0; i < $('#quranselection').get(0).options.length; i++) { $('#quranselection').get(0).options[i].text = $('#quranselection').get(0).options[i].text.replace(/(\(.*? )(.*?)(\u202C\))( )?(.*)/g, function (m, open, a, close) { return open + a + close + ' ' + (($('#translitscheme').val() === '0' || a === '') ? '' : (($('#translitscheme').val() !== '0' && $('#translitscheme').val() !== '1' && parseInt($('#translitscheme').val(), 10) % 2 !== 1) ? transliterateWithRules(a, parseInt($('#translitscheme').val(), 10) >= 2 ? Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2 : parseInt($('#translitscheme').val(), 10), null, false) : doTransliterate(a, true, parseInt($('#translitscheme').val(), 10)))); }); } }"}
        GetJS.AddRange(ArbDataWeb.GetUniCats())
        GetJS.AddRange(PlainTransliterateGenJS)
        GetJS.AddRange(TransliterateGenJS)
        GetJS.AddRange(NumberGenJS)
        Return GetJS.ToArray()
    End Function
    Public Function GetTransliterationSchemeTable(ByVal Item As PageLoader.TextItem, Context As HttpContext) As RenderArray
        Return New RenderArray("translitscheme") With {.Items = Arb.GetTransliterationTable(DecodeTranslitScheme(Context))}
    End Function
    Function GetTranslitSchemeJSArray() As String
        'Dim Letters(ArabicData.ArabicLetters.Length - 1) As IslamData.ArabicSymbol
        'ArabicData.ArabicLetters.CopyTo(Letters, 0)
        'Array.Sort(Letters, New StringLengthComparer("RomanTranslit"))
        Return "var translitSchemes = " + UtilityWeb.MakeJSIndexedObject(New List(Of String)(Linq.Enumerable.Select(ChData.IslamData.TranslitSchemes, Function(TranslitScheme As IslamData.TranslitScheme) CStr(Array.IndexOf(ChData.IslamData.TranslitSchemes, TranslitScheme) + 1))).ToArray(),
                                                                          New Array() {New List(Of String)(Linq.Enumerable.Select(ChData.IslamData.TranslitSchemes, Function(TranslitScheme As IslamData.TranslitScheme) UtilityWeb.MakeJSIndexedObject({"standard", "multi", "special", "gutteral"}, New Array() {New String() {UtilityWeb.MakeJSIndexedObject(New List(Of String)(Linq.Enumerable.Select(Arb.ArabicTranslitLetters(), Function(Str As String) System.Text.RegularExpressions.Regex.Replace(Str, "\(?\\u([0-9a-fA-F]{4})\)?", Function(Match As System.Text.RegularExpressions.Match) ChrW(Integer.Parse(Match.Groups(1).Value, Globalization.NumberStyles.HexNumber))))).ToArray(), New Array() {New List(Of String)(Linq.Enumerable.Select(Arb.ArabicTranslitLetters(),
                                                                            Function(Str As String) Arb.GetSchemeValueFromSymbol(ArbData.ArabicLetters(ArbData.FindLetterBySymbol(Str.Chars(0))), TranslitScheme.Name))).ToArray()}, False), UtilityWeb.MakeJSArray(TranslitScheme.Multis), UtilityWeb.MakeJSArray(TranslitScheme.SpecialLetters), UtilityWeb.MakeJSArray(TranslitScheme.Gutterals)}}, True))).ToArray()}, True)

    End Function
    Function GetArabicSymbolJSArray() As String
        GetArabicSymbolJSArray = "var arabicLetters = " +
                                UtilityWeb.MakeJSArray(New String() {UtilityWeb.MakeJSIndexedObject(New String() {"Symbol", "Shaping"},
                                New List(Of String())(Linq.Enumerable.Select(Array.FindAll(ArbData.ArabicLetters, Function(Letter As ArabicData.ArabicSymbol) Arb.GetSchemeValueFromSymbol(Letter, "ExtendedBuckwalter") <> String.Empty), Function(Convert As ArabicData.ArabicSymbol) New String() {CStr(AscW(Convert.Symbol)), If(Convert.Shaping = Nothing, String.Empty, UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(Convert.Shaping, Function(Ch As Char) CStr(AscW(Ch)))).ToArray()))})).ToArray(), False)}, True) + ";"
    End Function
    Public Shared FindLetterBySymbolJS As String = "function findLetterBySymbol(chVal) { var iSubCount; for (iSubCount = 0; iSubCount < arabicLetters.length; iSubCount++) { if (chVal === parseInt(arabicLetters[iSubCount].Symbol, 10)) return iSubCount; for (var iShapeCount = 0; iShapeCount < arabicLetters[iSubCount].Shaping.length; iShapeCount++) { if (chVal === parseInt(arabicLetters[iSubCount].Shaping[iShapeCount], 10)) return iSubCount; } } return -1; }"
    Public TransliterateGenJS As String()
    Public IsDiacriticJS As String
    Public Shared DiacriticJS As String() =
        {"function doScriptFormatChange(sVal, to, from) { return (to === 0 && from === 2) ? processTransform(sVal, simpleScriptBase.concat(simpleCleanScript), false) : sVal; }",
         "function processTransform(sVal, rules, bPriority) { var count, rep = []; for (count = 0; count < rules.length; count++) { if (bPriority) { sVal = sVal.replace(rules[count].match, function() { return (rules[count].negativematch !== '' && RegExp.matchResult(rules[count].evaluator, arguments[arguments.length - 2], arguments[arguments.length - 1], Array.prototype.slice.call(arguments).slice(0, -2)) ? arguments[0] : RegExp.matchResult(arguments[0], arguments[arguments.length - 2], arguments[arguments.length - 1], Array.prototype.slice.call(arguments).slice(0, -2))); }); } else { var arr, re = new RegExp(rules[count].match, 'g'); while ((arr = re.exec(sVal)) !== null) { if (rules[count].negativematch === '' || RegExp.matchResult(rules[count].negativematch, arr.index, sVal, arr) === '') { var dupCount; for (dupCount = 0; dupCount < RegExp.matchResult(rules[count].evaluator, arr.index, sVal, arr).length; dupCount++) { if (arr.index + dupCount >= sVal.length || sVal[arr.index + dupCount] !== RegExp.matchResult(rules[count].evaluator, arr.index, sVal, arr)[dupCount]) break; } rep.push({index: arr.index + dupCount, length: arr[0].length - dupCount, type: RegExp.matchResult(rules[count].evaluator, arr.index, sVal, arr).substr(dupCount), origOrder: count}); } } } } rep.sort(ruleMetadataComparer); for (count = 0; count < rep.length; count++) { sVal = sVal.substr(0, rep[count].index) + rep[count].type + sVal.substr(rep[count].index + rep[count].length); } return sVal; }"}
    Public PlainTransliterateGenJS As String()
    Public NumberGenJS As String()
    Public Function GetTransliterateNumberJS() As String()
        Dim GetJS As New List(Of String) From {"javascript: doTransliterateNum();", String.Empty, GetArabicSymbolJSArray(), GetTranslitSchemeJSArray()}
        GetJS.AddRange(ArbDataWeb.GetUniCats())
        GetJS.AddRange(PlainTransliterateGenJS)
        GetJS.AddRange(TransliterateGenJS)
        GetJS.AddRange(NumberGenJS)
        Return GetJS.ToArray()
    End Function
    Public Function GetTransliterateJS() As String()
        Dim GetJS As New List(Of String) From {"javascript: doTransliterateDisplay();", String.Empty, GetArabicSymbolJSArray(), GetTranslitSchemeJSArray(),
        "function doDirectionDom(elem, sVal, direction) { elem.css('direction', direction ? 'ltr' : 'rtl'); var stack = [], lastStrong = -1, lastCount = 0, iCount; for (iCount = 0; iCount < sVal.length; iCount++) { if (sVal.charCodeAt(iCount) === 0x200E || sVal.charCodeAt(iCount) === 0x200F || sVal.charCodeAt(iCount) === 0x61C) { if (lastStrong !== iCount - 1) {  } } else if (IsExplicit(sVal.charCodeAt(iCount))) { if (sVal.charCodeAt(iCount) === 0x202C || sVal.charCodeAt(iCount) === 0x2069) { stack.pop()[1].add(document.createTextNode(sVal.substring(lastCount, iCount - 1))); lastCount = iCount + 1; lastStrong = -1; } else { (stack.length === 0 ? elem : stack[stack.length - 1][1]).add(document.createTextNode(sVal.substring(lastCount, iCount - 1))); lastCount = iCount + 1; lastStrong = -1; stack.push([sVal[iCount], (stack.length === 0 ? elem : stack[stack.length - 1][1]).add('span')]); stack[stack.length - 1][1].css('direction', (sVal[iCount] === 0x202D || sVal[iCount] === 0x202A || sVal[iCount] === 0x2066) ? 'ltr' : 'rtl'); } } else if (!IsNeutral(sVal.charCodeAt(iCount))) { lastStrong = iCount; } } (stack.length === 0 ? elem : stack[stack.length - 1][1]).add(document.createTextNode(sVal.substring(lastCount, iCount - 1))); }",
        "function doTransliterateDisplay() { $('#translitvalue').css('direction', !$('#scheme1').prop('checked') && $($('#scheme0').prop('checked') ? '#direction0' : '#operation1').prop('checked') ? 'ltr' : 'rtl'); $('#translitvalue').empty(); $('#translitvalue').text($('#scheme0').prop('checked') ? doTransliterate($('#translitedit').val(), $('#direction0').prop('checked'), parseInt($('#translitscheme').val(), 10)) : ($('#scheme1').prop('checked') ? doScriptFormatChange($('#translitedit').val(), parseInt($('#toscheme').val(), 10), parseInt($('#fromscheme').val(), 10)) : $('#translitedit').val())); $('#translitvalue').html($('#translitvalue').html().replace(/\n/g, '<br>')); }"}
        GetJS.AddRange(ArbDataWeb.GetUniCats())
        GetJS.AddRange(PlainTransliterateGenJS)
        GetJS.AddRange(TransliterateGenJS)
        GetJS.AddRange(DiacriticJS)
        Return GetJS.ToArray()
    End Function
    Public Shared Function GetSchemeChangeJS() As String()
        Return New String() {"javascript: doSchemeChange();", String.Empty, "function doSchemeChange() { $('#wordbywordedit').css('display', $('#scheme3').prop('checked') ? 'block' : 'none'); $('#operation_').css('display', $('#scheme2').prop('checked') ? 'block' : 'none'); $('#fromscript_').css('display', $('#scheme1').prop('checked') ? 'block' : 'none'); $('#toscript_').css('display', $('#scheme1').prop('checked') ? 'block' : 'none'); $('#translitscheme_').css('display', $('#scheme0').prop('checked') ? 'block' : 'none'); $('#translitscheme').css('display', $('#scheme0').prop('checked') ? 'block' : 'none'); $('#direction_').css('display', $('#scheme0').prop('checked') ? 'block' : 'none'); }"}
    End Function
    Public Shared Function GetSchemeChangeUpdateJS() As String()
        Return New String() {"javascript: doSchemeChangeUpdate()", String.Empty, "function doSchemeChangeUpdate() { if (!$('#scheme3').prop('checked')) return; $('#translitvalue').attr('dir', 'rtl').css('height', '400px').css('overflow-y', 'scroll'); var srcs = doTransliterate($('#translitedit').val(), $('#direction0').prop('checked'), parseInt($('#translitscheme').val(), 10)).split(' '), tsrcs = $('#translitedit').val().split(' '), words = $('#wordbywordedit').val().split('|'), ctls = $('#translitvalue').children(), ct; for (ct = 0; ct < Math.max(srcs.length, words.length); ct++) { if (ct >= ctls.length) { $('#translitvalue').append(document.createElement('div')); ctls = $('#translitvalue').children(); ctls.eq(ct).addClass('multidisplay'); ctls.eq(ct).append(document.createElement('span')); ctls.eq(ct).append(document.createElement('span')); ctls.eq(ct).append(document.createElement('span')); ctls.eq(ct).children().eq(0).addClass('arabic'); ctls.eq(ct).children().eq(0).attr('dir', 'rtl'); ctls.eq(ct).children().eq(1).addClass('transliteration'); ctls.eq(ct).children().eq(1).attr('dir', 'ltr'); ctls.eq(ct).children().eq(2).addClass('translation'); ctls.eq(ct).children().eq(2).attr('dir', 'ltr'); } if ((ct < srcs.length ? srcs[ct] : '') !== ctls.eq(ct).children().eq(0).text()) ctls.eq(ct).children().eq(0).text((ct < srcs.length ? srcs[ct] : '')); if ((ct < tsrcs.length ? tsrcs[ct] : '') !== ctls.eq(ct).children().eq(1).text()) ctls.eq(ct).children().eq(1).text((ct < tsrcs.length ? tsrcs[ct] : '')); if ((ct < words.length ? words[ct] : '') !== ctls.eq(ct).children().eq(2).text()) ctls.eq(ct).children().eq(2).text((ct < words.length ? words[ct] : '')); } for (; ct < ctls.length; ct++) { ctls.eq(ct).remove(); } }"}
    End Function
End Class
Public Class ArabicFont
    Public _PortableMethods As PortableMethods
    Public ChData As CachedData
    Public Sub New(NewPortableMethods As PortableMethods, NewChData As CachedData)
        _PortableMethods = NewPortableMethods
        ChData = NewChData
    End Sub
    'Web.Config requires: configuration -> system.webServer -> staticContent -> <mimeMap fileExtension=".otf" mimeType="application/octet-stream" />
    'Web.Config requires for cross site scripting: configuration -> system.WebServer -> httpProtocol -> customHeaders -> <add name="Access-Control-Allow-Origin" value="*" />
    Public Function GetFontList() As Array()
        Dim Count As Integer
        Dim Strings(ChData.IslamData.ArabicFonts.Length - 1) As Array
        For Count = 0 To ChData.IslamData.ArabicFonts.Length - 1
            Strings(Count) = New String() {_PortableMethods.LoadResourceString("IslamInfo_" + ChData.IslamData.ArabicFonts(Count).Name), ChData.IslamData.ArabicFonts(Count).ID}
        Next
        Return Strings
    End Function
    Public Function GetArabicFontListJS() As String
        Return "var fontList = " +
        UtilityWeb.MakeJSIndexedObject(New List(Of String)(Linq.Enumerable.Select(ChData.IslamData.ArabicFonts, Function(Convert As IslamData.ArabicFontList) Convert.ID)).ToArray(), New Array() {New List(Of String)(Linq.Enumerable.Select(ChData.IslamData.ArabicFonts, Function(Convert As IslamData.ArabicFontList) UtilityWeb.MakeJSIndexedObject(New String() {"family", "embed", "file", "scale"}, New Array() {New String() {Convert.Family, Convert.EmbedName, Convert.FileName, CStr(Convert.Scale)}}, False))).ToArray()}, True) +
        ";var fontPrefs = " + UtilityWeb.MakeJSIndexedObject(New List(Of String)(Linq.Enumerable.Select(ChData.IslamData.ScriptFonts, Function(Convert As IslamData.ScriptFont) Convert.Name)).ToArray(),
                                                          New Array() {New List(Of String)(Linq.Enumerable.Select(ChData.IslamData.ScriptFonts, Function(Convert As IslamData.ScriptFont) UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(Of IslamData.ScriptFont.Font, String)(Convert.FontList, Function(SubConv As IslamData.ScriptFont.Font) SubConv.ID)).ToArray()))).ToArray()}, True) + ";"
    End Function
    Public Shared Function GetFontEmbedJS() As String
        Return "function embedFontStyle(fontID) { if (isInArray(embeddedFonts, fontID)) return; embeddedFonts.push(fontID); var font=fontList[fontID]; var style = 'font-family: \'' + font.embed + '\';' + 'src: url(\'/files/' + font.file + '.eot\');' + 'src: local(\'' + font.family + '\'), url(\'/files/' + font.file + ((font.file == 'KFC_naskh') ? '.otf\') format(\'opentype\');' : '.ttf\') format(\'truetype\');'); addStyleSheetRule(newStyleSheet(), '@font-face', style);  }"
    End Function
    Public Shared Function GetFontInitJS() As String
        Return "var tryFontCounter = 0; var embeddedFonts = " + UtilityWeb.MakeJSArray(New String() {"null"}, True) + "; var baseFont = 'Times New Roman';"
    End Function
    Public Shared Function GetFontIDJS() As String
        Return "function getFontID() { var fontID = $('#fontselection').val(); if (fontID == 'def') { fontID = 'me_quran'; if (isMac && isSafari) fontID = 'scheherazade'; if (isChrome) fontID = getPrefInstalledFont('uthmani'); } return fontID; }"
    End Function
    Public Shared Function GetFontFaceJS() As String
        Return "function getFontFace(fontID) { return fontList[fontID].family + (fontList[fontID].embed ? ',' + fontList[fontID].embed : ''); }"
    End Function
    Public Function GetFontWidthJS() As String
        Return "function fontWidth(fontName, text) { text = text || '" + UtilityWeb.EncodeJS(TanzilReader.GetTextVerse(TanzilReader.GetTextChapter(ChData.XMLDocMain, 3), 9).Attribute("text").Value) + "' ; if (text == 2) text = '" + UtilityWeb.EncodeJS(_PortableMethods.LoadResourceString("IslamInfo_InTheNameOfAllah")) + "," + UtilityWeb.EncodeJS(TanzilReader.GetTextVerse(TanzilReader.GetTextChapter(ChData.XMLDocMain, 1), 1).Attribute("text").Value) + "'; var tester = $('#font-tester'); tester.css('fontFamily', fontName); if (tester.firstChild) tester.remove(tester.firstChild); tester.append(document.createTextNode(text)); tester.css('display', 'block'); var width = tester.offsetWidth; tester.css('display', 'none'); return width; }"
    End Function
    Public Shared Function GetFontExistsJS() As String
        Return "function fontExists(fontName) { var fontFamily = fontName + ', ' + baseFont; return fontWidth(baseFont) * fontWidth(baseFont, 2) != fontWidth(fontFamily) * fontWidth(fontFamily, 2); }"
    End Function
    Public Shared Function GetApplyFontJS() As String
        Return "function applyFont(fontID) { if (!fontExists(getFontFace(fontID))) fontID = getPrefInstalledFont(); var font = fontList[fontID]; findStyleSheetRule('span.arabic').style.fontFamily = getFontFace(fontID); $('#fontloading').css('display', 'none'); }"
    End Function
    Public Shared Function GetTryFontJS() As String
        Return "function tryFont(fontID) { if (++tryFontCounter < 50 && !fontExists(getFontFace(fontID))) { setTimeout('tryFont(\'' + fontID + '\')', 400); return; } $('#fontloading').css('display', 'none'); applyFont(fontID); }"
    End Function
    Public Shared Function GetApplyEmbedFontJS() As String
        Return "function applyEmbedFont(fontID) { embedFontStyle(fontID); $('#fontloading').css('display', ''); tryFontCounter = 0; tryFont(fontID); }"
    End Function
    Public Shared Function GetFontPrefInstalledJS() As String
        Return "function getPrefInstalledFont(type) { var list = fontPrefs[type]; for(var i in list) { if (list.hasOwnProperty(i)) { var fontID = list[i]; if (fontList[fontID].installed) return fontID; } } return 'arial'; }"
    End Function
    Public Shared Function GetCheckInstalledFontsJS() As String
        Return "function checkInstalledFonts() { for (var i in fontList) { if (fontList.hasOwnProperty(i)) { var font = fontList[i]; if (font.family && fontExists(font.family)) font.installed = true; } } }"
    End Function
    Public Shared Function GetUpdateCustomFontJS() As String
        Return "function updateCustomFont() { var fontID = getFontID(); $('#fontcustom').css('display', fontID == 'custom' ? '' : 'none'); $('#fontcustomapply').css('display', fontID == 'custom' ? '' : 'none'); }"
    End Function
    Public Shared Function GetChangeCustomFontJS() As String()
        Return New String() {"javascript: changeCustomFont();", String.Empty, "function changeCustomFont() { fontList['custom'].family = $('#fontcustom').val(); fontList['custom'].scale = fontWidth(baseFont) / fontWidth(fontList['custom'].family); changeFont(); }"}
    End Function
    Public Function GetChangeFontJS() As String()
        Return New String() {"javascript: changeFont();", "checkInstalledFonts();", UtilityWeb.GetLookupStyleSheetJS(), GetArabicFontListJS(), UtilityWeb.GetBrowserTestJS(), UtilityWeb.GetAddStyleSheetJS(), UtilityWeb.GetAddStyleSheetRuleJS(), UtilityWeb.GetLookupStyleSheetJS(), UtilityWeb.IsInArrayJS(), GetUpdateCustomFontJS(), GetFontInitJS(), GetFontPrefInstalledJS(), GetCheckInstalledFontsJS(), GetFontIDJS(), GetFontFaceJS(), GetFontWidthJS(), GetFontExistsJS(), GetFontEmbedJS(), GetApplyFontJS(), GetTryFontJS(), GetApplyEmbedFontJS(),
        "function changeFont() { var fontID = getFontID(); updateCustomFont(); if (fontList[fontID].embed) applyEmbedFont(fontID); else applyFont(fontID); }"}
    End Function
    Public Shared Function GetFontSmallerJS() As String()
        Return New String() {"javascript: decreaseFontSize();", String.Empty, UtilityWeb.GetLookupStyleSheetJS(),
        "function decreaseFontSize() { rule = findStyleSheetRule('span.arabic'); rule.style.fontSize = Math.max(parseInt(rule.style.fontSize.replace('px', ''), 10) - 1, 1) + 'px'; $('.arabic > img').each(function (i) { this.src = this.src.replace(/Size=(\d+)/g, function (mat, p) { return 'Size=' + (parseInt(p) - 1).toString(); }); }); }"}
    End Function
    Public Shared Function GetFontDefaultSizeJS() As String()
        Return New String() {"javascript: defaultFontSize();", String.Empty, UtilityWeb.GetLookupStyleSheetJS(),
        "function defaultFontSize() { findStyleSheetRule('span.arabic').style.fontSize = '32px'; $('.arabic > img').each(function (i) { this.src = this.src.replace(/Size=(\d+)/g, 'Size=32'); }); }"}
    End Function
    Public Shared Function GetFontBiggerJS() As String()
        Return New String() {"javascript: increaseFontSize();", String.Empty, UtilityWeb.GetLookupStyleSheetJS(),
        "function increaseFontSize() { rule = findStyleSheetRule('span.arabic'); rule.style.fontSize = (parseInt(rule.style.fontSize.replace('px', ''), 10) + 1) + 'px'; $('.arabic > img').each(function (i) { this.src = this.src.replace(/Size=(\d+)/g, function (mat, p) { return 'Size=' + (parseInt(p) + 1).toString(); }); }); }"}
    End Function
End Class
Public Class CachedDataWeb
    Private ChData As CachedData
    Public Sub New(NewChData As CachedData)
        ChData = NewChData
    End Sub
    Public Function ArabicLetterCharacteristics(ByVal Item As PageLoader.TextItem) As String
        Return ChData.ArabicLetterCharacteristics()
    End Function
End Class
Public Class DocBuilderWeb
    Private _PortableMethods As PortableMethods
    Private ArbWeb As ArabicWeb
    Private ChData As CachedData
    Private Arb As Arabic
    Private ArbData As ArabicData
    Private TR As TanzilReader
    Private DocBuild As DocBuilder
    Private PhrWeb As PhrasesWeb
    Private TRWeb As TanzilReaderWeb
    Public Sub New(NewPortableMethods As PortableMethods, NewArbWeb As ArabicWeb, NewChData As CachedData, NewArb As Arabic, NewArbData As ArabicData, NewTR As TanzilReader, NewDocBuild As DocBuilder, NewPhrWeb As PhrasesWeb, NewTRWeb As TanzilReaderWeb)
        _PortableMethods = NewPortableMethods
        ArbWeb = NewArbWeb
        ChData = NewChData
        Arb = NewArb
        ArbData = NewArbData
        TR = NewTR
        DocBuild = NewDocBuild
        PhrWeb = NewPhrWeb
        TRWeb = NewTRWeb
    End Sub
    Public Async Function GetListRenderedText(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Task(Of RenderArray)
        Return Await DoGetListRenderedText(ArabicWeb.DecodeTranslitSchemeType(Context), ArbWeb.DecodeTranslitScheme(Context), CInt(Context.Request.QueryString.Get("selection")), Item.Name, Context.Request.QueryString.Get("qurantranslation"))
    End Function
    Public Async Function GetRenderedText(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Task(Of RenderArray)
        Return Await DoGetRenderedText(ArabicWeb.DecodeTranslitSchemeType(Context), ArbWeb.DecodeTranslitScheme(Context), Item.Name, Context.Request.Params("docedit"), Context.Request.Params("qurantranslation"))
    End Function
    Public Function GetMetadataRules(ID As String) As Array()
        Dim Output(ChData.IslamData.MetaRules.Length + 2) As Array
        Output(0) = New String() {}
        Output(1) = New String() {String.Empty, String.Empty, String.Empty}
        Output(2) = New String() {_PortableMethods.LoadResourceString("IslamInfo_Name"), _PortableMethods.LoadResourceString("IslamInfo_Translation"), _PortableMethods.LoadResourceString("IslamInfo_Translation")}
        For Count = 0 To ChData.RuleMetas("UthmaniQuran").Rules.Length - 1
            Output(3 + Count) = New Object() {ChData.RuleMetas("UthmaniQuran").Rules(Count).Name, New RenderArray.RenderItem() {New RenderArray.RenderItem(RenderArray.RenderTypes.eText, DocBuilder.ColorizeRegExGroups(DocBuilder.GetRegExText(ChData.RuleMetas("UthmaniQuran").Rules(Count).Match), False))}, New RenderArray.RenderItem() {New RenderArray.RenderItem(RenderArray.RenderTypes.eText, DocBuilder.ColorizeList(New List(Of String)(Linq.Enumerable.Select(ChData.RuleMetas("UthmaniQuran").Rules(Count).Evaluator, Function(Str As IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs()) DocBuilder.GetRegExText(String.Join("|"c, Linq.Enumerable.Select(Str, Function(S) S.RuleName + If(S.Args Is Nothing OrElse S.Args.Length = 0, "", "(" + String.Join(",", Linq.Enumerable.Select(S.Args, Function(Arg) String.Join(" ", Arg))) + ")")))))).ToArray(), False))}}
        Next
        Return RenderArrayWeb.MakeTableJSFunctions(Output, ID)
    End Function
    Public Function GetRuleSetRules(ID As String, Data As IslamData.RuleTranslationCategory.RuleTranslation()) As Array()
        Dim Output(Data.Length + 2) As Array
        Output(0) = New String() {}
        Output(1) = New String() {String.Empty, String.Empty, String.Empty}
        Output(2) = New String() {_PortableMethods.LoadResourceString("IslamInfo_Name"), _PortableMethods.LoadResourceString("IslamInfo_Translation"), _PortableMethods.LoadResourceString("IslamInfo_Translation")}
        For Count = 0 To Data.Length - 1
            Output(3 + Count) = New Object() {Data(Count).Name, New RenderArray.RenderItem() {New RenderArray.RenderItem(RenderArray.RenderTypes.eText, DocBuilder.ColorizeRegExGroups(DocBuilder.GetRegExText(Data(Count).Match), False))}, New RenderArray.RenderItem() {New RenderArray.RenderItem(RenderArray.RenderTypes.eText, DocBuilder.ColorizeRegExGroups(DocBuilder.GetRegExText(Data(Count).Evaluator), True))}}
        Next
        Return RenderArrayWeb.MakeTableJSFunctions(Output, ID)
    End Function
    Public Function GetRenderedHelpText(ByVal Item As PageLoader.TextItem, Context As HttpContext) As RenderArray
        Return DoGetRenderedHelpText(ArabicWeb.DecodeTranslitSchemeType(Context), ArbWeb.DecodeTranslitScheme(Context), Item.Name)
    End Function
    Public Function DoGetRenderedHelpText(SchemeType As ArabicData.TranslitScheme, Scheme As String, Name As String) As RenderArray
        Dim Renderer As New RenderArray(Name)
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, ArbWeb.GetTranslitSchemeMetadata("0"))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, TRWeb.GetTranslationMetadata("1"))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetMetadataRules("2"))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("3", ChData.RuleTranslations("RomanizationRules")))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("4", ChData.RuleTranslations("ColoringSpelledOutRules")))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("5", ChData.RuleTranslations("ErrorCheck")))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("6", ChData.RuleTranslations("WarshScript")))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("7", ChData.RuleTranslations("UthmaniMinimalScript")))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("8", ChData.RuleTranslations("SimpleEnhancedScript")))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("9", ChData.RuleTranslations("SimpleScript")))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("10", ChData.RuleTranslations("SimpleCleanScript")))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("11", ChData.RuleTranslations("SimpleMinimalScript")))}))
        Return Renderer
    End Function
    Public Async Function TextFromReferences(ID As String, Strings As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer) As Task(Of RenderArray)
        Dim _Options As String() = Strings.Split(";"c)
        Dim Count As Integer = 0
        While _Options(Count).EndsWith("&leftbrace") Or _Options(Count).EndsWith("&rightbrace") Or _Options(Count).EndsWith("&comma") Or _Options(Count).EndsWith("&semicolon")
            _Options(0) = _Options(0) + ";" + _Options(Count + 1)
            Count += 1
        End While
        Strings = _Options(0)
        Dim Options As New Dictionary(Of String, String())
        For Count = 1 To _Options.Length - 1
            Dim Vals As String() = _Options(Count).Split("="c)
            Options(Vals(0)) = If(Vals.Length = 1, Nothing, Vals(1).Split(","c))
        Next
        If Options.ContainsKey("Translation") Then TranslationIndex = TR.GetTranslationIndex(Options("Translation")(0))
        If Options.ContainsKey("TranslitStyle") Then
            If (Options("TranslitStyle")(0) = "Literal") Then
                SchemeType = ArabicData.TranslitScheme.Literal
            ElseIf (Options("TranslitStyle")(0) = "RuleBased") Then
                SchemeType = ArabicData.TranslitScheme.RuleBased
            ElseIf (Options("TranslitStyle")(0) = "LearningMode") Then
                SchemeType = ArabicData.TranslitScheme.LearningMode
            Else
                SchemeType = ArabicData.TranslitScheme.None
            End If
        End If
        If Options.ContainsKey("TranslitScheme") Then Scheme = Options("TranslitScheme")(0)
        Dim Renderer As New RenderArray(String.Empty)
        If Strings = Nothing Then Return Renderer
        'text before and after reference matches needs rendering
        'hadith reference matching {name,book/hadith}
        If TR.IsQuranTextReference(Strings) Then
            Renderer.Items.AddRange((Await TR.QuranTextFromReference(Strings, SchemeType, Scheme, TranslationIndex, Options.ContainsKey("W4W") Or Options.ContainsKey("W4WNum"), Options.ContainsKey("W4WNum"), Options.ContainsKey("NoArabic"), Options.ContainsKey("Header"), Options.ContainsKey("NoRef"), Options.ContainsKey("Colorize"), Options.ContainsKey("Verses"))).Items)
        ElseIf Strings.StartsWith("search:") Then
            Dim SelArr As String() = Strings.Replace("search:", String.Empty).Split(","c)
            Renderer.Items.AddRange((Await TR.QuranTextFromSearch(SelArr(0).Replace("&leftbrace;", "{").Replace("&rightbrace;", "}").Replace("&comma;", ",").Replace("&semicolon;", ";"), SchemeType, Scheme, TranslationIndex, Options.ContainsKey("W4W") Or Options.ContainsKey("W4WNum"), Options.ContainsKey("W4WNum"), Options.ContainsKey("NoArabic"), Options.ContainsKey("Header"), Options.ContainsKey("Colorize"), Options.ContainsKey("Verses"))).Items)
        ElseIf Strings.StartsWith("symbol:") Then
            Dim Symbols As New List(Of ArabicData.ArabicSymbol)
            Dim SelArr As String() = Strings.Replace("symbol:", String.Empty).Split(","c)
            For SubCount = 0 To ArbData.ArabicLetters.Length - 1
                If Array.IndexOf(SelArr, ArbData.ToCamelCase(ArbData.ArabicLetters(SubCount).UnicodeName).Replace("ArabicLetter", String.Empty).Replace("Arabic", String.Empty)) <> -1 Then
                    Symbols.Add(ArbData.ArabicLetters(SubCount))
                End If
            Next
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Arb.SymbolDisplay(Symbols.ToArray(), SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("personalpronoun:") Or Strings.StartsWith("proximaldemonstratives:") Or Strings.StartsWith("distaldemonstratives:") Then
            Dim Words As IslamData.GrammarSet.GrammarNoun()
            Dim SelArr As String()
            If Strings.StartsWith("proximaldemonstratives") Then
                Words = Arb.GetCatNoun("proxdemo")
                SelArr = Strings.Replace("proximaldemonstratives:", String.Empty).Split(","c)
            ElseIf Strings.StartsWith("distaldemonstratives") Then
                Words = Arb.GetCatNoun("distdemo")
                SelArr = Strings.Replace("distaldemonstratives:", String.Empty).Split(","c)
            ElseIf Strings.StartsWith("personalpronoun:") Then
                Words = Arb.GetCatNoun("perspro")
                SelArr = Strings.Replace("personalpronoun:", String.Empty).Split(","c)
            Else
                Words = Nothing
                SelArr = Nothing
            End If
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, ArbWeb.DisplayPronoun(Array.FindAll(Words, Function(Word As IslamData.GrammarSet.GrammarNoun) Array.IndexOf(SelArr, Word.TranslationID) <> -1), ID, True, SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("plurals:") Or Strings.StartsWith("possessivedeterminerpersonalpronoun:") Then
            Dim Words As IslamData.GrammarSet.GrammarTransform()
            Dim SelArr As String()
            If Strings.StartsWith("plurals") Then
                Words = Arb.GetTransform("mp|fp|bp")
                SelArr = Strings.Replace("plurals:", String.Empty).Split(","c)
            ElseIf Strings.StartsWith("possessivedeterminerpersonalpronoun") Then
                Words = Arb.GetTransform("posspron")
                SelArr = Strings.Replace("possessivedeterminerpersonalpronoun:", String.Empty).Split(","c)
            Else
                Words = Nothing
                SelArr = Nothing
            End If
            If Not Words Is Nothing Then Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, ArbWeb.DisplayTransform(String.Empty, Array.FindAll(Words, Function(Word As IslamData.GrammarSet.GrammarTransform) Array.IndexOf(SelArr, Word.TranslationID) <> -1), ID, True, True, SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("particle:") Then
            Dim SelArr As String() = Strings.Replace("particle:", String.Empty).Split(","c)
            Dim Words As New List(Of IslamData.GrammarSet.GrammarParticle)
            For Count = 0 To SelArr.Length - 1
                Words.AddRange(Arb.GetParticles(SelArr(Count)))
            Next
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, ArbWeb.DisplayParticle(Words.ToArray(), ID, SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("noun:") Then
            Dim SelArr As String() = Strings.Replace("noun:", String.Empty).Split(","c)
            Dim Words As New List(Of IslamData.GrammarSet.GrammarNoun)
            For Count = 0 To SelArr.Length - 1
                Words.AddRange(Arb.GetCatNoun(SelArr(Count)))
            Next
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, ArbWeb.NounDisplay(Words.ToArray(), ID, SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("verb:") Then
            Dim SelArr As String() = Strings.Replace("verb:", String.Empty).Split(","c)
            Dim Words As New List(Of IslamData.GrammarSet.GrammarVerb)
            For Count = 0 To SelArr.Length - 1
                Words.AddRange(Arb.GetVerb(SelArr(Count)))
            Next
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, ArbWeb.VerbDisplay(Words.ToArray(), ID, SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("transform:") Then
            Dim SelArr As String() = Strings.Replace("transform:", String.Empty).Split(","c)
            Dim Words As New List(Of IslamData.GrammarSet.GrammarTransform)
            For Count = 0 To SelArr.Length - 1
                Words.AddRange(Arb.GetTransform(SelArr(Count)))
            Next
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, ArbWeb.DisplayTransform(String.Empty, Words.ToArray(), ID, False, False, SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("word:") Then
            Dim Words As IslamData.GrammarSet.GrammarWord() = Arb.GetCatWords(Strings.Replace("word:", String.Empty).Split(","c))
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, ArbWeb.DisplayWord(Words, ID, SchemeType, Scheme))}))
        ElseIf Strings.StartsWith("phrase:") Then
            Dim SelArr As String() = Strings.Replace("phrase:", String.Empty).Split(","c)
            Dim PhraseCats As IslamData.Phrase() = Phrases.GetPhraseCats(SelArr, ChData.IslamData.Phrases)
            For Count = 0 To PhraseCats.Length - 1
                Renderer.Items.AddRange(Await PhrWeb.DoGetRenderedPhraseText(SchemeType, Scheme, PhraseCats(Count), TranslationIndex))
            Next
        ElseIf Strings.StartsWith("list:") Then
            Dim SelArr As String() = Strings.Replace("list:", String.Empty).Split(","c)
            Dim ListCats As IslamData.ListCategory.Word() = DocBuild.GetListCats(SelArr)
            For Count = 0 To ListCats.Length - 1
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, _PortableMethods.LoadResourceString("IslamInfo_" + ListCats(Count).TranslationID))}))
                Renderer.Items.AddRange((Await BuckwalterTextFromReferences(ID, SchemeType, Scheme, ListCats(Count).Text, String.Empty, TranslationIndex)).Items)
            Next
        ElseIf DocBuild.Abbrevs.ContainsKey(Strings) Then
            Dim AbbrevWord As IslamData.AbbrevWord = DocBuild.Abbrevs(Strings)
            Dim PhraseCat As Nullable(Of IslamData.Phrase) = Phrases.GetPhraseCat(AbbrevWord.TranslationID, ChData.IslamData.Phrases)
            Dim GrammarWord As Nullable(Of IslamData.GrammarSet.GrammarWord) = Arb.GetCatWord(AbbrevWord.TranslationID)
            Dim Items As New List(Of RenderArray.RenderItem)
            If AbbrevWord.Font <> String.Empty Then
                If Options.ContainsKey("Char") Then Options("Char") = New List(Of String)(Linq.Enumerable.Select(Options("Char"), Function(Str As String) Char.ConvertFromUtf32(Integer.Parse(Str, Globalization.NumberStyles.HexNumber)))).ToArray()
                If Options.Count = 0 Then Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTag, AbbrevWord.TranslationID + "|" + AbbrevWord.Text)}))
                For Each Str As String In AbbrevWord.Font.Split("|"c)
                    Dim Font As String = String.Empty
                    If Str.Contains(";") Then
                        Font = Str.Split(";"c)(0)
                        Str = Str.Split(";"c)(1)
                    End If
                    If Not Options.ContainsKey("Font") OrElse Array.IndexOf(Options("Font"), Font) <> -1 Then
                        For Each SubStr As String In Str.Split(","c)
                            If Not Options.ContainsKey("Char") OrElse Array.IndexOf(Options("Char"), String.Join(String.Empty, Linq.Enumerable.Select(SubStr.Split("+"c), Function(S As String) Char.ConvertFromUtf32(Integer.Parse(S, Globalization.NumberStyles.HexNumber))))) <> -1 Then
                                Dim RendText As New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, String.Join(String.Empty, Linq.Enumerable.Select(SubStr.Split("+"c), Function(Split As String) Char.ConvertFromUtf32(Integer.Parse(Split, System.Globalization.NumberStyles.HexNumber)))))
                                RendText.Font = Font
                                Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {RendText}))
                            End If
                        Next
                    End If
                Next
            End If
            If PhraseCat.HasValue Then
                Renderer.Items.AddRange(Await PhrWeb.DoGetRenderedPhraseText(SchemeType, Scheme, PhraseCat.Value, TranslationIndex))
                Renderer.Items.AddRange(Items)
            End If
            If GrammarWord.HasValue Then
                If AbbrevWord.Font <> String.Empty Then
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arb.TransliterateFromBuckwalter(GrammarWord.Value.Text)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arb.TransliterateToScheme(Arb.TransliterateFromBuckwalter(GrammarWord.Value.Text), SchemeType, Scheme, Arabic.FilterMetadataStops(Arb.TransliterateFromBuckwalter(GrammarWord.Value.Text), Arb.GetMetarules(Arb.TransliterateFromBuckwalter(GrammarWord.Value.Text), ChData.RuleMetas("Normal")), Nothing)).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, _PortableMethods.LoadResourceString("IslamInfo_" + GrammarWord.Value.TranslationID)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, Items)}))
                Else
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arb.TransliterateFromBuckwalter(GrammarWord.Value.Text)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arb.TransliterateToScheme(Arb.TransliterateFromBuckwalter(GrammarWord.Value.Text), SchemeType, Scheme, Arabic.FilterMetadataStops(Arb.TransliterateFromBuckwalter(GrammarWord.Value.Text), Arb.GetMetarules(Arb.TransliterateFromBuckwalter(GrammarWord.Value.Text), ChData.RuleMetas("Normal")), Nothing)).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, _PortableMethods.LoadResourceString("IslamInfo_" + GrammarWord.Value.TranslationID))}))
                End If
            End If
        ElseIf Strings.StartsWith("reference:") Then
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "(" + Strings.Replace("reference:", String.Empty) + ")")}))
        ElseIf Strings.StartsWith("arabic:") Then
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arb.TransliterateFromBuckwalter(Strings.Replace("arabic:", String.Empty).Replace("&leftbrace;", "{").Replace("&rightbrace;", "}").Replace("&comma;", ",").Replace("&semicolon;", ";")))}))
        ElseIf Strings.StartsWith("text:") Then
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Strings.Replace("text:", String.Empty))}))
        Else
        End If
        Return Renderer
    End Function
    Public Async Function BuckwalterTextFromReferences(ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, Strings As String, TranslationID As String, TranslationIndex As Integer) As Task(Of RenderArray)
        Dim Renderer As New RenderArray(ID)
        If Strings = Nothing Then Return Renderer
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Strings, "(.*?)(?:(\\\{)(.*?)(\\\})|$)", System.Text.RegularExpressions.RegexOptions.Singleline)
        Dim EnglishByWord As String() = If(TranslationID = Nothing, {}, (Await DocBuild.GetW4WItem(TranslationID)).Split("|"c))
        For MatchCount As Integer = 0 To Matches.Count - 1
            If Matches(MatchCount).Length <> 0 Then
                If Matches(MatchCount).Groups(1).Length <> 0 Then
                    Dim ArabicText As String() = Matches(MatchCount).Groups(1).Value.Split(" "c)
                    If ArabicText.Length > 1 Then 'And EnglishByWord.Length = ArabicText.Length Then
                        Dim Transliteration As String() = Arb.TransliterateToScheme(Arb.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), SchemeType, Scheme, Arabic.FilterMetadataStops(Arb.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), Arb.GetMetarules(Arb.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), ChData.RuleMetas("Normal")), Nothing)).Split(" "c)
                        'Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, _PortableMethods.LoadResourceString("IslamInfo_" + TranslationID))}))
                        Dim Items As New Collections.Generic.List(Of RenderArray.RenderItem)
                        For WordCount As Integer = 0 To EnglishByWord.Length - 1
                            Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arb.TransliterateFromBuckwalter(If(WordCount >= ArabicText.Length, String.Empty, ArabicText(WordCount)))), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(WordCount >= Transliteration.Length, String.Empty, Transliteration(WordCount))), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, EnglishByWord(WordCount))}))
                        Next
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, Items), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arb.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arb.TransliterateToScheme(Arb.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), SchemeType, Scheme, Arabic.FilterMetadataStops(Arb.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), Arb.GetMetarules(Arb.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), ChData.RuleMetas("Normal")), Nothing))), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, _PortableMethods.LoadResourceString("IslamInfo_" + TranslationID))}))
                    Else
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arb.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arb.TransliterateToScheme(Arb.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), SchemeType, Scheme, Arabic.FilterMetadataStops(Arb.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), Arb.GetMetarules(Arb.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), ChData.RuleMetas("Normal")), Nothing))), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, _PortableMethods.LoadResourceString("IslamInfo_" + TranslationID))}))
                    End If
                End If
                If Matches(MatchCount).Groups(3).Length <> 0 Then
                    Renderer.Items.AddRange((Await TextFromReferences(ID, Matches(MatchCount).Groups(3).Value, SchemeType, Scheme, TranslationIndex)).Items)
                End If
            End If
        Next
        Return Renderer
    End Function
    Public Async Function NormalTextFromReferences(ID As String, Strings As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer) As Task(Of RenderArray)
        Dim Renderer As New RenderArray(ID)
        If Strings = Nothing Then Return Renderer
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Strings, "(.*?)(?:(\{)(.*?)(\})|$)", System.Text.RegularExpressions.RegexOptions.Singleline)
        For Count As Integer = 0 To Matches.Count - 1
            If Matches(Count).Length <> 0 Then
                If Matches(Count).Groups(1).Length <> 0 Then
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.ePassThru, Matches(Count).Groups(1).Value)}))
                End If
                If Matches(Count).Groups(3).Length <> 0 Then
                    Renderer.Items.AddRange((Await TextFromReferences(ID + CStr(Renderer.Items.Count), Matches(Count).Groups(3).Value, SchemeType, Scheme, TranslationIndex)).Items)
                End If
            End If
        Next
        Return Renderer
    End Function
    Public Async Function DoGetListRenderedText(SchemeType As ArabicData.TranslitScheme, Scheme As String, Count As Integer, Name As String, QuranTranslation As String) As Task(Of RenderArray)
        If Count = -1 Then Count = 0
        Dim Renderer As New RenderArray(Name)
        If Not ChData.IslamData.Lists(Count).Words Is Nothing Then
            For SubCount = 0 To ChData.IslamData.Lists(Count).Words.Length - 1
                Renderer.Items.AddRange((Await BuckwalterTextFromReferences(Name, SchemeType, Scheme, ChData.IslamData.Lists(Count).Words(SubCount).Text, String.Empty, TR.GetTranslationIndex(QuranTranslation))).Items)
            Next
        End If
        Return Renderer
    End Function
    Public Async Function DoGetRenderedText(SchemeType As ArabicData.TranslitScheme, Scheme As String, Name As String, DocEdit As String, QuranTranslation As String) As Task(Of RenderArray)
        Return Await NormalTextFromReferences(Name, DocEdit, SchemeType, Scheme, TR.GetTranslationIndex(QuranTranslation))
    End Function
End Class
Public Class PhrasesWeb
    Private ArbWeb As ArabicWeb
    Private ChData As CachedData
    Private TR As TanzilReader
    Private DBWeb As DocBuilderWeb
    Public Sub New(NewArbWeb As ArabicWeb, NewChData As CachedData, NewTR As TanzilReader)
        ArbWeb = NewArbWeb
        ChData = NewChData
        TR = NewTR
    End Sub
    Public Sub Init(NewDBWeb As DocBuilderWeb)
        DBWeb = NewDBWeb
    End Sub
    Public Async Function GetRenderedPhraseText(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Task(Of RenderArray)
        Dim SchemeType As ArabicData.TranslitScheme = ArabicWeb.DecodeTranslitSchemeType(Context)
        Dim Scheme As String = ArbWeb.DecodeTranslitScheme(Context)
        Return Await DoGetRenderedCatText(Item.Name, SchemeType, Scheme, ChData.IslamData.Phrases, TR.GetTranslationIndex(Context.Request.QueryString.Get("qurantranslation")))
    End Function
    Public Async Function DoGetRenderedPhraseText(SchemeType As ArabicData.TranslitScheme, Scheme As String, Verse As IslamData.Phrase, TranslationIndex As Integer) As Task(Of List(Of RenderArray.RenderItem))
        Return (Await DBWeb.BuckwalterTextFromReferences(String.Empty, SchemeType, Scheme, Verse.Text, Verse.TranslationID, TranslationIndex)).Items
    End Function
    Public Async Function DoGetRenderedCatText(ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, Category As IslamData.Phrase(), TranslationIndex As Integer) As Task(Of RenderArray)
        Dim Renderer As New RenderArray(ID)
        For SubCount As Integer = 0 To Category.Length - 1
            Renderer.Items.AddRange(Await DoGetRenderedPhraseText(SchemeType, Scheme, Category(SubCount), TranslationIndex))
        Next
        Return Renderer
    End Function
End Class
Public Class Quiz
    Private Arb As Arabic
    Private ArbWeb As ArabicWeb
    Private ChData As CachedData
    Private ArbDataWeb As ArabicDataWeb
    Public Sub New(NewArb As Arabic, NewArbWeb As ArabicWeb, NewChData As CachedData, NewArbDataWeb As ArabicDataWeb)
        Arb = NewArb
        ArbWeb = NewArbWeb
        ChData = NewChData
        ArbDataWeb = NewArbDataWeb
    End Sub
    Public Shared Function GetQuizList() As Array()
        Dim Strings(2) As Array
        Strings(0) = New String() {"ArabicLetters", "arabicletters"}
        Strings(1) = New String() {"ArabicDiacriticsLetters", "arabicdiacriticsletters"}
        Strings(2) = New String() {"ArabicLettersDiacritics", "arabiclettersdiacritics"}
        Return Strings
    End Function
    Public Shared ArabicSpecialLetters As String() = {ArabicData.ArabicLetterAlefWithHamzaAbove, ArabicData.ArabicLetterAlefWithHamzaBelow, ArabicData.ArabicLetterWawWithHamzaAbove, ArabicData.ArabicLetterYehWithHamzaAbove, ArabicData.ArabicLetterAlefWasla, ArabicData.ArabicLetterAlefWithMaddaAbove, ArabicData.ArabicLetterHamza, ArabicData.ArabicLetterTehMarbuta, ArabicData.ArabicLetterAlefMaksura}
    Public Shared ArabicDiacriticsBefore As String() = {ArabicData.ArabicSukun, ArabicData.ArabicFatha, ArabicData.ArabicKasra, ArabicData.ArabicDamma}
    Public Shared ArabicDiacriticsAfter As String() = {ArabicData.ArabicSukun, ArabicData.ArabicFatha, ArabicData.ArabicKasra, ArabicData.ArabicDamma, ArabicData.ArabicKasratan, ArabicData.ArabicDammatan, ArabicData.ArabicFathatan + ArabicData.ArabicLetterAlef, ArabicData.ArabicShadda}
    Public Shared Function GetChangeQuizJS() As String()
        Return New String() {"javascript: changeQuiz();", String.Empty,
                             "function changeQuiz() { qtype = $('#quizselection').val(); qwrong = 0; qright = 0; nextQuestion(); }"}
    End Function
    Public Shared Function DisplayCount(ByVal Item As PageLoader.TextItem) As String
        Return "Wrong: 0 Right: 0"
    End Function
    Public Function GetQuizSet(Context As HttpContext) As String()
        Dim Quiz As Integer = CInt(Context.Request.QueryString.Get("quizselection"))
        If Quiz = 0 Then
            Return ChData.ArabicLetters
        ElseIf Quiz = 1 Then
            Dim CurList As New Generic.List(Of String)
            For Count As Integer = 0 To ArabicDiacriticsBefore.Length - 1
                Dim CurLet As String = ArabicDiacriticsBefore(Count)
                CurList.AddRange(Linq.Enumerable.Select(ChData.ArabicLetters, Function(Str As String) CurLet + Str))
            Next
            Return CurList.ToArray()
        ElseIf Quiz = 2 Then
            Dim CurList As New Generic.List(Of String)
            For Count As Integer = 0 To ArabicDiacriticsAfter.Length - 1
                Dim CurLet As String = ArabicDiacriticsAfter(Count)
                CurList.AddRange(Linq.Enumerable.Select(ChData.ArabicLetters, Function(Str As String) Str + CurLet))
            Next
            Return CurList.ToArray()
        Else
            Return Nothing
        End If
    End Function
    Public Function DisplayQuestion(ByVal Item As PageLoader.TextItem, Context As HttpContext) As String
        Context.Items.Add("rnd", Guid.NewGuid().GetHashCode())
        Dim Rd As New Random(CInt(Context.Items("rnd")))
        Dim Count As Integer = CInt(Math.Floor(Rd.Next() * 4))
        Dim QuizSet As New List(Of String)
        QuizSet.AddRange(GetQuizSet(Context))
        While Count <> 0
            QuizSet.RemoveAt(CInt(Math.Floor(Rd.Next() * QuizSet.Count)))
            Count -= 1
        End While
        Return QuizSet(CInt(Math.Floor(Rd.Next() * QuizSet.Count)))
    End Function
    Public Function DisplayAnswer(ByVal Item As PageLoader.ButtonItem, Context As HttpContext) As String
        Dim Count As Integer
        Dim Rd As New Random(CInt(Context.Items("rnd")))
        Rd.Next()
        Dim Quiz As Integer = CInt(Context.Request.QueryString.Get("quizselection"))
        Dim QuizSet As New List(Of String)
        QuizSet.AddRange(GetQuizSet(Context))
        For Count = 2 To Integer.Parse(Item.Name.Replace("answer", String.Empty))
            QuizSet.RemoveAt(CInt(Math.Floor(Rd.Next() * QuizSet.Count)))
        Next
        Dim SchemeType As ArabicData.TranslitScheme = ArabicWeb.DecodeTranslitSchemeType(Context)
        Dim Scheme As String = ArbWeb.DecodeTranslitScheme(Context)
        If SchemeType = ArabicData.TranslitScheme.None Then
            SchemeType = ArabicData.TranslitScheme.LearningMode
            Scheme = ChData.IslamData.TranslitSchemes(3).Name
        End If
        Return Arb.TransliterateToScheme(QuizSet(CInt(Math.Floor(Rd.Next() * QuizSet.Count))), SchemeType, Scheme, Arabic.FilterMetadataStops(QuizSet(CInt(Math.Floor(Rd.Next() * QuizSet.Count))), Arb.GetMetarules(QuizSet(CInt(Math.Floor(Rd.Next() * QuizSet.Count))), ChData.RuleMetas("Normal")), Nothing))
    End Function
    Public Function VerifyAnswer() As String()
        Dim JSList As New List(Of String) From {"javascript: verifyAnswer(this);", String.Empty,
            ArbWeb.GetArabicSymbolJSArray(), ArbWeb.GetTranslitSchemeJSArray(), ArabicWeb.FindLetterBySymbolJS,
            "var arabicLets = " + UtilityWeb.MakeJSArray(ChData.ArabicLetters, False) + ";",
            "var arabicDiacriticsBefore = " + UtilityWeb.MakeJSArray(ArabicDiacriticsBefore, False) + ";",
            "var arabicDiacriticsAfter = " + UtilityWeb.MakeJSArray(ArabicDiacriticsAfter, False) + ";",
            "var qtype = 'arabicletters', qwrong = 0, qright = 0; if (typeof renderList == 'undefined') { renderList = []; }",
            "function getUniqueRnd(excl, count) { var rnd; if (excl.length === count) return 0; do { rnd = Math.floor(Math.random() * count); } while (excl.indexOf(rnd) !== -1); return rnd; }",
            "function getQuizSet() { if (qtype === 'arabicletters') return arabicLets; if (qtype === 'arabiclettersdiacritics') { var count = 0, arr = []; for (count = 0; count < arabicDiacriticsAfter.length; count++) { arr = arr.concat(arabicLets.map(function(val) { return val + arabicDiacriticsAfter[count]; })); } return arr; } if (qtype === 'arabicdiacriticsletters') { var count = 0, arr = []; for (count = 0; count < arabicDiacriticsBefore.length; count++) { arr = arr.concat(arabicLets.map(function(val) { return arabicDiacriticsBefore[count] + val; })); } return arr; }; return []; }",
            "function getQA(quizSet, quest, nidx) { if (quest) return quizSet[nidx]; return (parseInt($('#translitscheme').val(), 10) % 2) === 0 ? transliterateWithRules(quizSet[nidx], parseInt($('#translitscheme').val(), 10) >= 2 ? Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2 : 5, null, false) : doTransliterate(quizSet[nidx], true, parseInt($('#translitscheme').val(), 10)); }",
            "function nextQuestion() { $('#count').text('Wrong: ' + qwrong + ' Right: ' + qright); var i = Math.floor(Math.random() * 4), quizSet = getQuizSet(), pos = quizSet.length, nidx = getUniqueRnd([], pos), aidx = []; aidx[0] = getUniqueRnd([nidx], pos); aidx[1] = getUniqueRnd([nidx, aidx[0]], pos); aidx[2] = getUniqueRnd([nidx, aidx[0], aidx[1]], pos); $('#quizquestion').text(getQA(quizSet, true, nidx)); $('#answer1').prop('value', getQA(quizSet, false, i === 0 ? nidx : aidx[0])); $('#answer2').prop('value', getQA(quizSet, false, i === 1 ? nidx : aidx[i > 1 ? 1 : 0])); $('#answer3').prop('value', getQA(quizSet, false, i === 2 ? nidx : aidx[i > 2 ? 2 : 1])); $('#answer4').prop('value', getQA(quizSet, false, i === 3 ? nidx : aidx[2])); }",
            "function verifyAnswer(ctl) { $(ctl).prop('value') === ((parseInt($('#translitscheme').val(), 10) % 2) === 0 ? transliterateWithRules($('#quizquestion').text().trim(), parseInt($('#translitscheme').val(), 10) >= 2 ? Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2 : 5, null, false) : doTransliterate($('#quizquestion').text().trim(), true, parseInt($('#translitscheme').val(), 10))) ? qright++ : qwrong++; nextQuestion(); }"}
        JSList.AddRange(ArbDataWeb.GetUniCats())
        JSList.AddRange(ArbWeb.PlainTransliterateGenJS)
        JSList.AddRange(ArbWeb.TransliterateGenJS)
        Return JSList.ToArray()
    End Function
End Class
Public Class TanzilReaderWeb
    Private _PortableMethods As PortableMethods
    Private ChData As CachedData
    Private Arb As Arabic
    Private ArbWeb As ArabicWeb
    Private TR As TanzilReader
    Public Sub New(NewPortableMethods As PortableMethods, NewChData As CachedData, NewArb As Arabic, NewArbWeb As ArabicWeb, NewTR As TanzilReader)
        _PortableMethods = NewPortableMethods
        ChData = NewChData
        Arb = NewArb
        ArbWeb = NewArbWeb
        TR = NewTR
    End Sub
    Public Shared Function GetScriptFormatChangeJS() As String()
        Return New String() {"javascript: doScriptFormatOptChange(this);", String.Empty, "function doScriptFormatOptChange(obj) { var ct, oth = $(obj.id === 'toscript' ? '#fromscript' : '#toscript')[0]; for (ct = 0; ct < oth.options.length - 1 - 1; ct++) { if (obj.options[ct].value !== oth.options[(ct >= oth.selectedIndex) ? ct + 1 : ct].value) break; } oth.options.add(new Option(obj.options[ct].text, obj.options[ct].value), (ct >= oth.selectedIndex) ? ct + 1 : ct); for (ct = 0; ct < oth.options.length - 1; ct++) { if (oth.options[ct].value === obj.options[obj.selectedIndex].value) { oth.options.remove(ct); } } }"}
    End Function
    Public Function GetTranslationMetadata(ID As String) As Array()
        Dim Output(ChData.IslamData.Translations.TranslationList.Length + 2) As Array
        Output(0) = New String() {}
        Output(1) = New String() {String.Empty, String.Empty}
        Output(2) = New String() {_PortableMethods.LoadResourceString("IslamInfo_Name"), _PortableMethods.LoadResourceString("IslamInfo_Translation")}
        For Count = 0 To ChData.IslamData.Translations.TranslationList.Length - 1
            Output(3 + Count) = {ChData.IslamData.Translations.TranslationList(Count).FileName, ChData.IslamData.Translations.TranslationList(Count).Translator + " (" + ChData.IslamData.Translations.TranslationList(Count).Name + ")"}
        Next
        Return RenderArrayWeb.MakeTableJSFunctions(Output, ID)
    End Function
    Public Function GetDivisionChangeJS(Context As HttpContext) As String()
        Dim JSArrays As String = UtilityWeb.MakeJSArray(New String() {UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(GetChapterNames(Context), Function(Convert As Array) UtilityWeb.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))}))).ToArray(), True),
            UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(GetChapterNamesByRevelationOrder(Context), Function(Convert As Array) UtilityWeb.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))}))).ToArray(), True),
            UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(GetPartNames(Context), Function(Convert As Array) UtilityWeb.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))}))).ToArray(), True),
            UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(TR.GetGroupNames(), Function(Convert As Array) UtilityWeb.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))}))).ToArray(), True),
            UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(TR.GetStationNames(), Function(Convert As Array) UtilityWeb.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))}))).ToArray(), True),
            UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(TR.GetSectionNames(), Function(Convert As Array) UtilityWeb.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))}))).ToArray(), True),
            UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(TR.GetPageNames(), Function(Convert As Array) UtilityWeb.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))}))).ToArray(), True),
            UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(TR.GetSajdaNames(), Function(Convert As Array) UtilityWeb.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))}))).ToArray(), True),
            UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(TR.GetImportantNames(), Function(Convert As Array) UtilityWeb.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))}))).ToArray(), True),
            UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(Arb.GetRecitationSymbols(ArabicData.TranslitScheme.None, String.Empty), Function(Convert As Array) UtilityWeb.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))}))).ToArray(), True),
            UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(TR.GetRecitationRules(), Function(Convert As Array) UtilityWeb.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))}))).ToArray(), True)}, True)
        Return New String() {"javascript: changeQuranDivision(this.selectedIndex);", String.Empty, UtilityWeb.GetClearOptionListJS(),
        "function changeQuranDivision(index) { var iCount; var qurandata = " + JSArrays + "; var eSelect = $('#quranselection').get(0); clearOptionList(eSelect); for (iCount = 0; iCount < qurandata[index].length; iCount++) { eSelect.options.add(new Option(qurandata[index][iCount][0], qurandata[index][iCount][1])); } }"}
    End Function
    Public Shared Function GetWordVerseModeJS() As String()
        Return New String() {"javascript: changeWordVerseMode(this.selectedIndex);", String.Empty,
                             "function changeWordVerseMode(index) {}"}
    End Function
    Public Shared Function GetColorCueModeJS() As String()
        Return New String() {"javascript: changeColorCueMode(this.selectedIndex);", String.Empty,
                             "function changeColorCueMode(index) {}"}
    End Function
    Public Function GetQuranWordTotal(ByVal Item As PageLoader.TextItem, Context As HttpContext) As String
        Return TR.GetQuranWordTotal(Context.Request.QueryString.Get("quranselection"))
    End Function
    Public Async Function GetQuranWordFrequency(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Task(Of Object())
        Return Await TR.GetQuranWordFrequency(ArabicWeb.DecodeTranslitSchemeType(Context), ArbWeb.DecodeTranslitScheme(Context), Context.Request.QueryString.Get("quranselection"))
    End Function
    Public Function GetSelectionNames(Context As HttpContext) As Array()
        Return TR.GetSelectionNames(Context.Request.QueryString.Get("qurandivision"), ArabicWeb.DecodeTranslitSchemeType(Context), ArbWeb.DecodeTranslitScheme(Context))
    End Function
    Public Async Function GetRenderedQuranText(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Task(Of RenderArray)
        Return Await TR.GetRenderedQuranText(ArabicWeb.DecodeTranslitSchemeType(Context), ArbWeb.DecodeTranslitScheme(Context), Item.Name, Context.Request.QueryString.Get("qurandivision"), Context.Request.QueryString.Get("quranselection"), Context.Request.QueryString.Get("qurantranslation"), Context.Request.QueryString.Get("wordversemode"), Context.Request.QueryString.Get("colorcuemode"))
    End Function
    Public Function GetChapterNames(Context As HttpContext) As Array()
        Return TR.GetChapterNames(ArabicWeb.DecodeTranslitSchemeType(Context), ArbWeb.DecodeTranslitScheme(Context))
    End Function
    Public Function GetChapterNamesByRevelationOrder(Context As HttpContext) As Array()
        Return TR.GetChapterNamesByRevelationOrder(ArabicWeb.DecodeTranslitSchemeType(Context), ArbWeb.DecodeTranslitScheme(Context))
    End Function
    Public Function GetPartNames(Context As HttpContext) As Array()
        Return TR.GetPartNames(ArabicWeb.DecodeTranslitSchemeType(Context), ArbWeb.DecodeTranslitScheme(Context))
    End Function
End Class
Public Class HadithReaderWeb
    Private _PortableMethods As PortableMethods
    Private ChData As CachedData
    Private Arb As Arabic
    Private ArbWeb As ArabicWeb
    Private HR As HadithReader
    Private UWeb As UtilityWeb
    Private ISDatabaseLookup As IslamSiteDatabaseLookup
    Public Sub New(NewPortableMethods As PortableMethods, NewArb As Arabic, NewChData As CachedData, NewArbWeb As ArabicWeb, NewHR As HadithReader, NewUWeb As UtilityWeb, NewISDatabaseLookup As IslamSiteDatabaseLookup)
        _PortableMethods = NewPortableMethods
        Arb = NewArb
        ChData = NewChData
        ArbWeb = NewArbWeb
        HR = NewHR
        UWeb = NewUWeb
        ISDatabaseLookup = NewISDatabaseLookup
    End Sub
    Public Function GetCollectionChangeOnlyJS() As String
        Dim JSArrays As String = UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(ChData.IslamData.Collections, Function(Convert As IslamData.CollectionInfo) UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(Of IslamData.CollectionInfo.CollTranslationInfo, String)(Convert.Translations, Function(TranslateBlock As IslamData.CollectionInfo.CollTranslationInfo) UtilityWeb.MakeJSArray(New String() {_PortableMethods.LoadResourceString("lang_local" + Languages.GetLanguageInfoByCode(TranslateBlock.FileName.Substring(0, 2), ChData.IslamData.LanguageList).Code) + ": " + _PortableMethods.LoadResourceString("IslamInfo_" + TranslateBlock.Name), TranslateBlock.FileName}))).ToArray(), True))).ToArray(), True)
        Return "function changeHadithCollection(index) { var iCount; var hadithdata = " + JSArrays + "; var eSelect = $('#hadithtranslation').get(0); clearOptionList(eSelect); for (iCount = 0; iCount < hadithdata[index].length; iCount++) { eSelect.options.add(new Option(hadithdata[index][iCount][0], hadithdata[index][iCount][1])); } }"
    End Function
    Public Function GetCollectionChangeWithBooksJS() As String()
        Dim JSArrays As String = UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(Of IslamData.CollectionInfo, String)(ChData.IslamData.Collections, Function(Convert As IslamData.CollectionInfo) UtilityWeb.MakeJSArray(New List(Of String)(Linq.Enumerable.Select(Of Array, String)(HR.GetBookNamesByCollection(HR.GetCollectionIndex(Convert.Name)), Function(BookNames As Array) UtilityWeb.MakeJSArray(New String() {CStr(BookNames.GetValue(0)), CStr(BookNames.GetValue(1))}))).ToArray(), True))).ToArray(), True)
        Return New String() {"javascript: changeHadithCollectionBooks(this.selectedIndex);", String.Empty, UtilityWeb.GetClearOptionListJS(),
                             GetCollectionChangeOnlyJS(),
        "function changeHadithCollectionBooks(index) { changeHadithCollection(index); var iCount; var hadithtdata = " + JSArrays + "; var eSelect = $('#hadithbook').get(0); clearOptionList(eSelect); for (iCount = 0; iCount < hadithtdata[index].length; iCount++) { eSelect.options.add(new Option(hadithtdata[index][iCount][0], hadithtdata[index][iCount][1])); } }"}
    End Function
    Public Function GetCollectionChangeJS() As String()
        Return New String() {"javascript: changeHadithCollection(this.selectedIndex);", String.Empty, UtilityWeb.GetClearOptionListJS(),
                             GetCollectionChangeOnlyJS()}
    End Function
    Public Function GetCollectionXMLMetaDataDownload(Context As HttpContext) As String()
        Return New String() {UWeb.GetPageString("Source&File=" + ChData.IslamData.Collections(GetCurrentCollection(Context)).FileName + "-data.xml"), _PortableMethods.LoadResourceString("IslamInfo_" + ChData.IslamData.Collections(GetCurrentCollection(Context)).Name) + " XML metadata"}
    End Function
    Public Function GetCollectionXMLDownload(Context As HttpContext) As String()
        Return New String() {UWeb.GetPageString("Source&File=" + ChData.IslamData.Collections(GetCurrentCollection(Context)).FileName + ".xml"), _PortableMethods.LoadResourceString("IslamInfo_" + ChData.IslamData.Collections(GetCurrentCollection(Context)).Name) + " XML source text"}
    End Function
    Public Function GetTranslationXMLMetaDataDownload(Context As HttpContext) As String()
        Dim TranslationIndex As Integer = HR.GetTranslationIndex(GetCurrentCollection(Context), Context.Request.QueryString.Get("hadithtranslation"))
        If TranslationIndex = -1 Then Return New String() {}
        Return New String() {UWeb.GetPageString("Source&File=" + HR.GetTranslationXMLFileName(GetCurrentCollection(Context), Context.Request.QueryString.Get("hadithtranslation")) + ".xml"), ChData.IslamData.Collections(GetCurrentCollection(Context)).Translations(TranslationIndex).Name + " XML metadata"}
    End Function
    Public Function GetTranslationTextDownload(Context As HttpContext) As String()
        Dim TranslationIndex As Integer = HR.GetTranslationIndex(GetCurrentCollection(Context), Context.Request.QueryString.Get("hadithtranslation"))
        If TranslationIndex = -1 Then Return New String() {}
        Return New String() {UWeb.GetPageString("Source&File=" + HR.GetTranslationFileName(GetCurrentCollection(Context), Context.Request.QueryString.Get("hadithtranslation")) + ".txt"), ChData.IslamData.Collections(GetCurrentCollection(Context)).Translations(TranslationIndex).Name + " raw source text"}
    End Function
    Public Shared Function GetCurrentCollection(Context As HttpContext) As Integer
        Dim Strings As String = Context.Request.QueryString.Get("hadithcollection")
        If Not Strings Is Nothing Then Return CInt(Strings) Else Return 0
    End Function
    Public Shared Function GetCurrentBook(Context As HttpContext) As Integer
        Dim Strings As String = Context.Request.QueryString.Get("hadithbook")
        If Not Strings Is Nothing Then Return CInt(Strings) Else Return 1
    End Function
    Public Function GetBookNames(Context As HttpContext) As Array()
        Return HR.GetBookNamesByCollection(GetCurrentCollection(Context))
    End Function
    Public Function GetTranslationList(Context As HttpContext) As Array()
        Return HR.GetTranslationList(GetCurrentCollection(Context))
    End Function
    Public Async Function GetHadithMappingText(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Task(Of Array())
        Return Await HR.GetHadithMappingText(GetCurrentCollection(Context), Context.Request.QueryString.Get("hadithtranslation"))
    End Function
    Public Async Function GetRenderedText(ByVal Item As PageLoader.TextItem, Context As HttpContext) As Task(Of RenderArray)
        Dim SchemeType As ArabicData.TranslitScheme = ArabicWeb.DecodeTranslitSchemeType(Context)
        Dim Scheme As String = ArbWeb.DecodeTranslitScheme(Context)
        Dim Translation As String = Context.Request.QueryString.Get("hadithtranslation")
        Return Await DoGetRenderedText(Item.Name, SchemeType, Scheme, Translation, GetCurrentCollection(Context), GetCurrentBook(Context), ISDatabaseLookup, Context)
    End Function
    Public Async Function DoGetRenderedText(ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, Translation As String, Index As Integer, BookIndex As Integer, RankInterface As IslamSiteRankingData, Context As HttpContext) As Threading.Tasks.Task(Of RenderArray)
        Dim Renderer As New RenderArray(ID)
        Dim Hadith As Integer
        Dim HadithText As Collections.Generic.List(Of Collections.Generic.List(Of Object)) = Await HR.GetHadithText(BookIndex, Index)
        Dim ChapterIndex As Integer = -1
        Dim SubChapterIndex As Integer = -1
        Dim BookNode As Xml.Linq.XElement = HR.GetBookByIndex(Index, BookIndex)
        Dim ChapterNode As Xml.Linq.XElement = Nothing
        Dim SubChapterNode As Xml.Linq.XElement
        If Not BookNode Is Nothing Then
            Dim Str As String = Arb.TransliterateFromBuckwalter("Had~ivu " + BookNode.Attribute("hadiths").Value + " ")
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Str), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arb.TransliterateToScheme(Str, SchemeType, Scheme, Arabic.FilterMetadataStops(Str, Arb.GetMetarules(Str, ChData.RuleMetas("Normal")), Nothing)).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + BookNode.Attribute("hadiths").Value + " ")}))
            Str = Arb.TransliterateFromBuckwalter("{lokita`bu " + CStr(BookIndex)) + " " + BookNode.Attribute("name").Value + " "
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Str), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arb.TransliterateToScheme(Str, SchemeType, Scheme, Arabic.FilterMetadataStops(Str, Arb.GetMetarules(Str, ChData.RuleMetas("Normal")), Nothing)).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Book " + CStr(BookIndex) + ": " + HR.GetBookEName(BookNode, Index) + " ")}))
            'Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderRight, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("mjld " + New List(Of Xml.Linq.XElement)(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocInfo(Index).Root.Elements).ToArray()).Elements).ToArray()(BookIndex).Attribute("volume").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("mjld " + New List(Of Xml.Linq.XElement)(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocInfo(Index).Root.Elements).ToArray()).Elements).ToArray()(BookIndex).Attribute("volume").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Volume " + Utility.GetChildNode("books", XMLDocInfo(Index).Root.Nodes).Nodes.Item(BookIndex).Attribute("volume").Value + " ")}))
            Dim XMLDocTranslate As Xml.Linq.XDocument = Nothing
            Dim Strings() As String = Nothing
            If ChData.IslamData.Collections(Index).Translations.Length <> 0 Then
                Dim Stream As IO.Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", HR.GetTranslationXMLFileName(Index, Translation) + ".xml")))
                XMLDocTranslate = Xml.Linq.XDocument.Load(Stream)
                Stream.Dispose()
                Strings = Await _PortableMethods.ReadAllLines(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", HR.GetTranslationFileName(Index, Translation) + ".txt")))
            End If
            For Hadith = 0 To HadithText.Count - 1
                'Handle missing or excess chapter indexes
                If ChapterIndex <> CInt(HadithText(Hadith)(1)) Then
                    ChapterIndex = CInt(HadithText(Hadith)(1))
                    ChapterNode = HadithReader.GetChapterByIndex(BookNode, ChapterIndex)
                    If Not ChapterNode Is Nothing Then
                        Str = Arb.TransliterateFromBuckwalter("Had~ivu " + ChapterNode.Attribute("hadiths").Value + " ")
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Str), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arb.TransliterateToScheme(Str, SchemeType, Scheme, Arabic.FilterMetadataStops(Str, Arabic.FilterMetadataStops(Str, Arb.GetMetarules(Str, ChData.RuleMetas("Normal")), Nothing), Nothing)).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + ChapterNode.Attribute("hadiths").Value + " ")}))
                        Dim Heads As New List(Of RenderArray.RenderText)
                        Heads.AddRange(Linq.Enumerable.Select(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split(Arb.TransliterateFromBuckwalter("baAbu " + CStr(ChapterIndex)) + " " + ChapterNode.Attribute("name").Value + " ", "(\d+\.\d+(?:-\d+)?)"), Function(S As String) If(System.Text.RegularExpressions.Regex.Match(S, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + S.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", S}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, S))))
                        Str = Arb.TransliterateFromBuckwalter("baAbu " + CStr(ChapterIndex)) + " " + System.Text.RegularExpressions.Regex.Replace(ChapterNode.Attribute("name").Value, "(\d+).(\d+(?:-\d+)?)", String.Empty) + " "
                        Heads.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arb.TransliterateToScheme(Str, SchemeType, Scheme, Arabic.FilterMetadataStops(Str, Arabic.FilterMetadataStops(Str, Arb.GetMetarules(Str, ChData.RuleMetas("Normal")), Nothing), Nothing)).Trim()))
                        Heads.AddRange(Linq.Enumerable.Select(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split("Chapter " + CStr(ChapterIndex) + ": " + Utility.DefaultValue(_PortableMethods.LoadResourceString("IslamInfo_" + ChData.IslamData.Collections(Index).FileName + "Book" + BookNode.Attribute("index").Value + "Chapter" + ChapterNode.Attribute("index").Value), String.Empty) + " ", "(\d+\.\d+(?:-\d+)?)"), Function(S As String) If(System.Text.RegularExpressions.Regex.Match(S, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + S.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", S}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, S))))
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, Heads.ToArray()))
                    End If
                    SubChapterIndex = -1
                End If
                'Handle missing or excess subchapter indexes
                If SubChapterIndex <> CInt(HadithText(Hadith)(2)) Then
                    SubChapterIndex = CInt(HadithText(Hadith)(2))
                    If Not ChapterNode Is Nothing Then
                        SubChapterNode = HadithReader.GetSubChapterByIndex(ChapterNode, SubChapterIndex)
                        If Not SubChapterNode Is Nothing Then
                            Str = Arb.TransliterateFromBuckwalter("Had~ivu " + SubChapterNode.Attribute("hadiths").Value + " ")
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arb.TransliterateFromBuckwalter("Had~ivu " + SubChapterNode.Attribute("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arb.TransliterateToScheme(Str, SchemeType, Scheme, Arb.GetMetarules(Str, ChData.RuleMetas("Normal"))).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + SubChapterNode.Attribute("hadiths").Value + " ")}))
                            Dim Heads As New List(Of RenderArray.RenderText)
                            Heads.AddRange(Linq.Enumerable.Select(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split(Arb.TransliterateFromBuckwalter("baAbu " + CStr(SubChapterIndex)) + " " + SubChapterNode.Attribute("name").Value + " ", "(\d+\.\d+(?:-\d+)?)"), Function(S As String) If(System.Text.RegularExpressions.Regex.Match(S, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + S.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", S}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, S))))
                            Str = Arb.TransliterateFromBuckwalter("baAbu " + CStr(SubChapterIndex)) + " " + System.Text.RegularExpressions.Regex.Replace(SubChapterNode.Attribute("name").Value, "(\d+).(\d+(?:-\d+)?)", String.Empty) + " "
                            Heads.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arb.TransliterateToScheme(Str, SchemeType, Scheme, Arb.GetMetarules(Str, ChData.RuleMetas("Normal"))).Trim()))
                            Heads.AddRange(Linq.Enumerable.Select(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split("Sub-Chapter " + CStr(SubChapterIndex) + ": " + Utility.DefaultValue(_PortableMethods.LoadResourceString("IslamInfo_" + ChData.IslamData.Collections(Index).FileName + "Book" + BookNode.Attribute("index").Value + "Chapter" + ChapterNode.Attribute("index").Value + "Subchapter" + SubChapterNode.Attribute("index").Value), String.Empty) + " ", "(\d+\.\d+(?:-\d+)?)"), Function(S As String) If(System.Text.RegularExpressions.Regex.Match(S, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + S.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", S}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, S))))
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, Heads.ToArray()))
                        End If
                    End If
                End If
                Dim HadithTranslation As String = String.Empty
                If CInt(HadithText(Hadith)(0)) <> 0 Then
                    Dim TranslationLines() As String = HR.GetTranslationHadith(XMLDocTranslate, Strings, Index, BookIndex - 1, CInt(HadithText(Hadith)(0)))
                    Dim Count As Integer
                    For Count = 0 To TranslationLines.Length - 1
                        HadithTranslation += vbCrLf + TranslationLines(Count)
                    Next
                End If
                'Arabic.TransliterateFromBuckwalter("(" + HadithText(Hadith)(0) + ") ")
                Dim Texts As New List(Of RenderArray.RenderText)
                Texts.AddRange(Linq.Enumerable.Select(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split(CStr(HadithText(Hadith)(3)) + " " + Arb.TransliterateFromBuckwalter("=" + CStr(HadithText(Hadith)(0)) + " "), "(\d+\.\d+(?:-\d+)?)"), Function(S As String) If(System.Text.RegularExpressions.Regex.Match(S, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + S.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", S}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, S))))
                Str = System.Text.RegularExpressions.Regex.Replace(CStr(HadithText(Hadith)(3)), "(\d+).(\d+(?:-\d+)?)", String.Empty) + " " + Arb.TransliterateFromBuckwalter("=" + CStr(HadithText(Hadith)(0))) + " "
                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arb.TransliterateToScheme(Str, SchemeType, Scheme, Arabic.FilterMetadataStops(Str, Arb.GetMetarules(Str, ChData.RuleMetas("Normal")), Nothing)).Trim()))
                Texts.AddRange(Linq.Enumerable.Select(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split("(" + CStr(HadithText(Hadith)(0)) + ") " + HadithTranslation, "(\d+\.\d+(?:-\d+)?)"), Function(S As String) If(System.Text.RegularExpressions.Regex.Match(S, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + S.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", S}), New RenderArray.RenderText(If(HR.IsTranslationTextLTR(Index, Translation), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), S))))
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                Dim Ranking As Integer() = RankInterface.GetHadithRankingData(ChData.IslamData.Collections(Index).FileName, BookIndex, CInt(HadithText(Hadith)(0)))
                Dim UserRanking As Integer = RankInterface.GetUserHadithRankingData(ChData.IslamData.Collections(Index).FileName, BookIndex, CInt(HadithText(Hadith)(0)), Context)
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eInteractive, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eRanking, ChData.IslamData.Collections(Index).FileName + "|" + CStr(BookIndex) + "|" + CStr(HadithText(Hadith)(0)) + "|" + CStr(Ranking(0)) + "|" + CStr(Ranking(1)) + "|" + CStr(UserRanking))}))
            Next
        End If
        Return Renderer
    End Function
End Class
Public Interface IslamSiteRankingData
    Function GetHadithRankingData(ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer) As Integer()
    Function GetUserHadithRankingData(ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer, Context As HttpContext) As Integer
End Interface
Public Class IslamSiteDatabaseLookup
    Implements IslamSiteRankingData
    Private ISDatabase As IslamSiteDatabase
    Private UWeb As UtilityWeb
    Public Sub New(NewUWeb As UtilityWeb, NewISDatabase As IslamSiteDatabase)
        UWeb = NewUWeb
        ISDatabase = NewISDatabase
    End Sub
    Public Function GetHadithRankingData(Collection As String, Book As Integer, Hadith As Integer) As Integer() Implements IslamSiteRankingData.GetHadithRankingData
        Return ISDatabase.GetHadithRankingData(Collection, Book, Hadith)
    End Function
    Public Function GetUserHadithRankingData(Collection As String, Book As Integer, Hadith As Integer, Context As HttpContext) As Integer Implements IslamSiteRankingData.GetUserHadithRankingData
        If Not UWeb.IsLoggedIn(Context) Then Return -1
        Dim UserID As Integer = UWeb.GetUserID(Context)
        Return ISDatabase.GetUserHadithRankingData(UserID, Collection, Book, Hadith)
    End Function
End Class

Public Class IslamSiteDatabase
    Private SD As SiteDatabase
    Public Sub New(NewSD As SiteDatabase)
        SD = NewSD
    End Sub
    Public Sub CreateDatabase()
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SD.GetConnection()
        If Connection Is Nothing Then Return
        SiteDatabase.ExecuteNonQuery(Connection, "CREATE TABLE HadithRankings (UserID int NOT NULL, " +
        "Collection VARCHAR(254) NOT NULL, " +
        "BookIndex int, " +
        "HadithIndex int NOT NULL, " +
        "Ranking int NOT NULL)")
        Connection.Close()
    End Sub
    Public Sub RemoveDatabase()
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SD.GetConnection()
        If Connection Is Nothing Then Return
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        SiteDatabase.ExecuteNonQuery(Connection, "DROP TABLE HadithRankings")
        Connection.Close()
    End Sub
    Public Sub UpdateRankingData(ByVal UserID As Integer, Context As HttpContext)
        If GetUserHadithRankingData(UserID, Context.Request.Form.Get("Collection"), CInt(Context.Request.Form.Get("Book")), CInt(Context.Request.Form.Get("Hadith"))) = -1 Then
            If CInt(Context.Request.Form.Get("Rating")) <> 0 Then
                SetUserHadithRankingData(UserID, Context.Request.Form.Get("Collection"), CInt(Context.Request.Form.Get("Book")), CInt(Context.Request.Form.Get("Hadith")), CInt(Context.Request.Form.Get("Rating")))
            End If
        Else
            If CInt(Context.Request.Form.Get("Rating")) <> 0 Then
                UpdateUserHadithRankingData(UserID, Context.Request.Form.Get("Collection"), CInt(Context.Request.Form.Get("Book")), CInt(Context.Request.Form.Get("Hadith")), CInt(Context.Request.Form.Get("Rating")))
            Else
                RemoveUserHadithRankingData(UserID, Context.Request.Form.Get("Collection"), CInt(Context.Request.Form.Get("Book")), CInt(Context.Request.Form.Get("Hadith")))
            End If
        End If
    End Sub
    Public Sub WriteRankingData(Context As HttpContext)
        Dim Data As Integer() = GetHadithRankingData(Context.Request.Form.Get("Collection"), CInt(Context.Request.Form.Get("Book")), CInt(Context.Request.Form.Get("Hadith")))
        If Data(1) <> 0 Then Context.Response.Write("Average of " + CStr(Data(0) / Data(1) / 2) + " out of " + CStr(Data(1)) + " rankings")
    End Sub
    Public Function GetHadithCollectionRankingData(ByVal Collection As String) As Double
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SD.GetConnection()
        If Connection Is Nothing Then Return -1
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT AVG(Ranking) AS AvgRank FROM HadithRankings Collection=@Collection"
        Command.Parameters.AddWithValue("@Collection", Collection)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetHadithCollectionRankingData = Reader.GetInt32("AvgRank")
        Else
            GetHadithCollectionRankingData = -1
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Function GetHadithBookRankingData(ByVal Collection As String, ByVal Book As Integer) As Double
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SD.GetConnection()
        If Connection Is Nothing Then Return -1
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT AVG(Ranking) AS AvgRank FROM HadithRankings WHERE Collection=@Collection AND BookIndex=" + CStr(Book)
        Command.Parameters.AddWithValue("@Collection", Collection)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetHadithBookRankingData = Reader.GetInt32("AvgRank")
        Else
            GetHadithBookRankingData = -1
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Function GetHadithRankingData(ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer) As Integer()
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SD.GetConnection()
        If Connection Is Nothing Then Return {0, 0}
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT SUM(Ranking) AS SumRank, COUNT(Ranking) AS CountRank FROM HadithRankings WHERE Collection=@Collection AND BookIndex=" + CStr(Book) + " AND HadithIndex=" + CStr(Hadith)
        Command.Parameters.AddWithValue("@Collection", Collection)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetHadithRankingData = {Reader.GetInt32("SumRank"), Reader.GetInt32("CountRank")}
        Else
            GetHadithRankingData = {0, 0}
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Function GetUserHadithRankingData(UserID As Integer, ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer) As Integer
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SD.GetConnection()
        If Connection Is Nothing Then Return -1
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT Ranking FROM HadithRankings WHERE UserID=" + CStr(UserID) + " AND Collection=@Collection AND BookIndex=" + CStr(Book) + " AND HadithIndex=" + CStr(Hadith)
        Command.Parameters.AddWithValue("@Collection", Collection)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetUserHadithRankingData = Reader.GetInt32("Ranking")
        Else
            GetUserHadithRankingData = -1
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Sub SetUserHadithRankingData(ByVal UserID As Integer, ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer, ByVal Rank As Integer)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SD.GetConnection()
        If Connection Is Nothing Then Return
        SiteDatabase.ExecuteNonQuery(Connection, "INSERT INTO HadithRankings (UserID, Collection, BookIndex, HadithIndex, Ranking) VALUES (" + CStr(UserID) + ", @Collection, " + CStr(Book) + ", " + CStr(Hadith) + ", " + CStr(Rank) + ")", New Generic.Dictionary(Of String, Object) From {{"@Collection", Collection}})
        Connection.Close()
    End Sub
    Public Sub UpdateUserHadithRankingData(ByVal UserID As Integer, ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer, ByVal Rank As Integer)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SD.GetConnection()
        If Connection Is Nothing Then Return
        SiteDatabase.ExecuteNonQuery(Connection, "UPDATE HadithRankings SET Ranking=" + CStr(Rank) + " WHERE UserID=" + CStr(UserID) + " AND Collection=@Collection AND BookIndex=" + CStr(Book) + " AND HadithIndex=" + CStr(Hadith), New Generic.Dictionary(Of String, Object) From {{"@Collection", Collection}})
        Connection.Close()
    End Sub
    Public Sub RemoveUserHadithRankingData(ByVal UserID As Integer, ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SD.GetConnection()
        If Connection Is Nothing Then Return
        SiteDatabase.ExecuteNonQuery(Connection, "DELETE FROM HadithRankings WHERE UserID=" + CStr(UserID) + " AND Collection=@Collection AND BookIndex=" + CStr(Book) + " AND HadithIndex=" + CStr(Hadith), New Generic.Dictionary(Of String, Object) From {{"@Collection", Collection}})
        Connection.Close()
    End Sub
End Class