<?php
/*
Plugin Name: IslamSource
Description: Islam Source Quranic Verse and Islamic Phrase Plugin - Allows for Quranic chapters, verses even specified down through the word to be inserted easily by using formats {a:b:c-x:y:z} where b, c, x, y and z are optional depending on if a chapter, verse or particular word of a verse is desired or a range is desired so it could be in forms {a:b-y} or {a:b-x:y} such as the opening chapter which could be specified as {1:1-7}.  The Arabic is automatically displayed when posts are viewed.  It also allows for various calligraphy and Unicode Islamic words or phrases to be easily inserted through a button on the visual editor which are displayed when the posts are later viewed.
Version: 1.0.0
Author: IslamSource
Author URI: http://islamsource.info
*/
class Utility
{
	public static $_resxml = null;
	public static function LoadResourceString($resourcekey)
	{
		if (Utility::$_resxml === null) Utility::$_resxml = simplexml_load_file(dirname(__FILE__) . "/IslamResources/MyProject/Resources.resx");
		$arr = Utility::$_resxml->xpath("/root/data[@name='" . $resourcekey . "']");
		if (count($arr) !== 0) return (string)$arr[0]->value;
		return null;
	}
	public static function GetChildNodeByIndex($nodename, $indexname, $index, $childnodes)
	{
		$xmlnode = $childnodes[$index];
		if ($index - 1 < count($childnodes)) {
			$xmlnode = $childnodes[$index - 1];
			if ($xmlnode !== null && $xmlnode->getName() == $nodename) {
				if (isset($xmlnode->attributes()[$indexname]) && (string)$xmlnode->attributes()[$indexname] == $index) {
					return $xmlnode;
				}
			}
		}
		foreach ($childnodes as $xmlnode) {
			if ($xmlnode->getName() == $nodename) {
				if (isset($xmlnode->attributes()[$indexname]) && (string)$xmlnode->attributes()[$indexname] == $index) {
					return $xmlnode;
				}				
			}
		}
		return null;
	}
	public static function GetChildNodeCount($nodename, $node)
	{
		$count = 0;
		for ($index = 0; $index < count($node->children()); $index++) {
            if ($node->children()[$index]->getName() == $nodename) $count += 1;
        }
        return $count;
    }
}

class Arabic
{
    Shared _ArabicXMLData As ArabicXMLData
    Public Shared ReadOnly Property Data As ArabicXMLData
        Get
            If _ArabicXMLData Is Nothing Then
                Dim fs As IO.FileStream = New IO.FileStream(Utility.GetFilePath(If(Utility.IsDesktopApp(), "HostPage\metadata\arabicdata.xml", "metadata\arabicdata.xml")), IO.FileMode.Open, IO.FileAccess.Read)
                Dim xs As System.Xml.Serialization.XmlSerializer = New System.Xml.Serialization.XmlSerializer(GetType(ArabicXMLData))
                _ArabicXMLData = CType(xs.Deserialize(fs), ArabicXMLData)
                fs.Close()
            End If
            Return _ArabicXMLData
        End Get
    End Property
	Public Shared _BuckwalterMap As Dictionary(Of Char, Integer)
    Public Shared ReadOnly Property BuckwalterMap As Dictionary(Of Char, Integer)
        Get
            If _BuckwalterMap Is Nothing Then
                _BuckwalterMap = New Dictionary(Of Char, Integer)
                For Index = 0 To Data.ArabicLetters.Length - 1
                    If Data.ArabicLetters(Index).ExtendedBuckwalterLetter <> ChrW(0) Then
                        _BuckwalterMap.Add(Data.ArabicLetters(Index).ExtendedBuckwalterLetter, Index)
                    End If
                Next
            End If
            Return _BuckwalterMap
        End Get
    End Property
    Public Shared Function TransliterateFromBuckwalter(ByVal Buckwalter As String) As String
        Dim ArabicString As String = String.Empty
        Dim Count As Integer
        For Count = 0 To Buckwalter.Length - 1
            If Buckwalter(Count) = "\" Then
                Count += 1
                If Buckwalter(Count) = "," Then
                    ArabicString += ArabicData.ArabicComma
                Else
                    ArabicString += Buckwalter(Count)
                End If
            Else
                If BuckwalterMap.ContainsKey(Buckwalter(Count)) Then
                    ArabicString += Data.ArabicLetters(BuckwalterMap.Item(Buckwalter(Count))).Symbol
                Else
                    ArabicString += Buckwalter(Count)
                End If
            End If
        Next
        Return ArabicString
    End Function
    Public Shared _ArabicLetterMap As Dictionary(Of Char, Integer)
    Public Shared ReadOnly Property ArabicLetterMap As Dictionary(Of Char, Integer)
        Get
            If _ArabicLetterMap Is Nothing Then
                _ArabicLetterMap = New Dictionary(Of Char, Integer)
                For Index = 0 To Data.ArabicLetters.Length - 1
                    If Data.ArabicLetters(Index).Symbol <> ChrW(0) Then
                        _ArabicLetterMap.Add(Data.ArabicLetters(Index).Symbol, Index)
                    End If
                Next
            End If
            Return _ArabicLetterMap
        End Get
    End Property
    Public Shared Function FindLetterBySymbol(Symbol As Char) As Integer
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
    Public Const LeftToRightOverride As Char = ChrW(&H202D)
    Public Const RightToLeftOverride As Char = ChrW(&H202E)
    Public Const NarrowNoBreakSpace As Char = ChrW(&H202F)
    Public Const DottedCircle As Char = ChrW(&H25CC)
    Public Const OrnateLeftParenthesis As Char = ChrW(&HFD3E)
    Public Const OrnateRightParenthesis As Char = ChrW(&HFD3F)
    Public Shared Function MakeUniRegEx(Input As String) As String
        Return String.Join(String.Empty, Array.ConvertAll(Of Char, String)(Input.ToCharArray(), Function(Ch As Char) "\u" + AscW(Ch).ToString("X4")))
    End Function
    Public Shared Function MakeRegMultiEx(Input As String()) As String
        Return String.Join("|", Input)
    End Function
    Public Enum TranslitScheme
        None = 0
        Literal = 1
        RuleBased = 2
    End Enum
};

class CachedData
{
	public static $_IslamData = null;
	public static function IslamData() { if (CachedData::$_IslamData === null) CachedData::$_IslamData = simplexml_load_file(dirname(__FILE__) . "/metadata/islaminfo.xml"); return CachedData::$_IslamData; }
    Shared _AlDuri As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _Warsh As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _WarshScript As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _UthmaniMinimalScript As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _SimpleEnhancedScript As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _SimpleScript As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _SimpleCleanScript As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _SimpleMinimalScript As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _RomanizationRules As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _ColoringSpelledOutRules As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _ErrorCheck As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _RulesOfRecitationRegEx As IslamData.RuleMetadataTranslation()
    Public Shared Function GetNum(Name As String) As String()
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.ArabicNumbers.Length - 1
            If CachedData.IslamData.ArabicNumbers(Count).Name = Name Then
                Return CachedData.IslamData.ArabicNumbers(Count).Text
            End If
        Next
        Return {}
    End Function
    Public Shared Function GetPattern(Name As String) As String
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.ArabicPatterns.Length - 1
            If CachedData.IslamData.ArabicPatterns(Count).Name = Name Then
                Return TranslateRegEx(CachedData.IslamData.ArabicPatterns(Count).Match, True)
            End If
        Next
        Return String.Empty
    End Function
    Public Shared Function GetGroup(Name As String) As String()
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.ArabicGroups.Length - 1
            If CachedData.IslamData.ArabicGroups(Count).Name = Name Then
                Return Array.ConvertAll(CachedData.IslamData.ArabicGroups(Count).Text, Function(Str As String) TranslateRegEx(Str, Name = "ArabicSpecialLetters"))
            End If
        Next
        Return {}
    End Function
    Public Shared Function GetRuleSet(Name As String) As IslamData.RuleTranslationCategory.RuleTranslation()
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.RuleSets.Length - 1
            If CachedData.IslamData.RuleSets(Count).Name = Name Then
                Dim RuleSet As IslamData.RuleTranslationCategory.RuleTranslation() = CachedData.IslamData.RuleSets(Count).Rules
                For SubCount As Integer = 0 To RuleSet.Length - 1
                    RuleSet(SubCount).Match = TranslateRegEx(RuleSet(SubCount).Match, True)
                    RuleSet(SubCount).Evaluator = TranslateRegEx(RuleSet(SubCount).Evaluator, False)
                Next
                Return RuleSet
            End If
        Next
        Return Nothing
    End Function
    Shared _ArabicUniqueLetters As String()
    Shared _ArabicNumbers As String()
    Shared _ArabicWaslKasraExceptions As String()
    Shared _ArabicBaseNumbers As String()
    Shared _ArabicBaseExtraNumbers As String()
    Shared _ArabicBaseTenNumbers As String()
    Shared _ArabicBaseHundredNumbers As String()
    Shared _ArabicBaseThousandNumbers As String()
    Shared _ArabicBaseMillionNumbers As String()
    Shared _ArabicBaseMilliardNumbers As String()
    Shared _ArabicBaseBillionNumbers As String()
    Shared _ArabicBaseTrillionNumbers As String()
    Shared _ArabicFractionNumbers As String()
    Shared _ArabicOrdinalNumbers As String()
    Shared _ArabicOrdinalExtraNumbers As String()
    Public Shared ReadOnly Property ArabicUniqueLetters As String()
        Get
            If _ArabicUniqueLetters Is Nothing Then
                _ArabicUniqueLetters = GetNum("ArabicUniqueLetters")
            End If
            Return _ArabicUniqueLetters
        End Get
    End Property
    Public Shared ReadOnly Property ArabicNumbers As String()
        Get
            If _ArabicNumbers Is Nothing Then
                _ArabicNumbers = GetNum("ArabicNumbers")
            End If
            Return _ArabicNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicWaslKasraExceptions As String()
        Get
            If _ArabicWaslKasraExceptions Is Nothing Then
                _ArabicWaslKasraExceptions = GetNum("ArabicWaslKasraExceptions")
            End If
            Return _ArabicWaslKasraExceptions
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseNumbers As String()
        Get
            If _ArabicBaseNumbers Is Nothing Then
                _ArabicBaseNumbers = GetNum("base")
            End If
            Return _ArabicBaseNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseExtraNumbers As String()
        Get
            If _ArabicBaseExtraNumbers Is Nothing Then
                _ArabicBaseExtraNumbers = GetNum("baseextras")
            End If
            Return _ArabicBaseExtraNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseTenNumbers As String()
        Get
            If _ArabicBaseTenNumbers Is Nothing Then
                _ArabicBaseTenNumbers = GetNum("baseten")
            End If
            Return _ArabicBaseTenNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseHundredNumbers As String()
        Get
            If _ArabicBaseHundredNumbers Is Nothing Then
                _ArabicBaseHundredNumbers = GetNum("basehundred")
            End If
            Return _ArabicBaseHundredNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseThousandNumbers As String()
        Get
            If _ArabicBaseThousandNumbers Is Nothing Then
                _ArabicBaseThousandNumbers = GetNum("thousands")
            End If
            Return _ArabicBaseThousandNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseMillionNumbers As String()
        Get
            If _ArabicBaseMillionNumbers Is Nothing Then
                _ArabicBaseMillionNumbers = GetNum("millions")
            End If
            Return _ArabicBaseMillionNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseMilliardNumbers As String()
        Get
            If _ArabicBaseMilliardNumbers Is Nothing Then
                _ArabicBaseMilliardNumbers = GetNum("milliard")
            End If
            Return _ArabicBaseMilliardNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseBillionNumbers As String()
        Get
            If _ArabicBaseBillionNumbers Is Nothing Then
                _ArabicBaseBillionNumbers = GetNum("billions")
            End If
            Return _ArabicBaseBillionNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseTrillionNumbers As String()
        Get
            If _ArabicBaseTrillionNumbers Is Nothing Then
                _ArabicBaseTrillionNumbers = GetNum("trillions")
            End If
            Return _ArabicBaseTrillionNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicFractionNumbers As String()
        Get
            If _ArabicFractionNumbers Is Nothing Then
                _ArabicFractionNumbers = GetNum("fractions")
            End If
            Return _ArabicFractionNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicOrdinalNumbers As String()
        Get
            If _ArabicOrdinalNumbers Is Nothing Then
                _ArabicOrdinalNumbers = GetNum("ordinals")
            End If
            Return _ArabicOrdinalNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicOrdinalExtraNumbers As String()
        Get
            If _ArabicOrdinalExtraNumbers Is Nothing Then
                _ArabicOrdinalExtraNumbers = GetNum("ordinalextras")
            End If
            Return _ArabicOrdinalExtraNumbers
        End Get
    End Property
    Shared _CertainStopPattern As String
    Shared _OptionalStopPattern As String
    Shared _OptionalStopPatternNotEndOfAyah As String
    Shared _CertainNotStopPattern As String
    Shared _OptionalNotStopPattern As String
    Shared _TehMarbutaStopRule As String
    Shared _TehMarbutaContinueRule As String
    Public Shared ReadOnly Property CertainStopPattern As String
        Get
            If _CertainStopPattern Is Nothing Then
                _CertainStopPattern = GetPattern("CertainStopPattern")
            End If
            Return _CertainStopPattern
        End Get
    End Property
    Public Shared ReadOnly Property OptionalStopPattern As String
        Get
            If _OptionalStopPattern Is Nothing Then
                _OptionalStopPattern = GetPattern("OptionalStopPattern")
            End If
            Return _OptionalStopPattern
        End Get
    End Property
    Public Shared ReadOnly Property OptionalStopPatternNotEndOfAyah As String
        Get
            If _OptionalStopPatternNotEndOfAyah Is Nothing Then
                _OptionalStopPatternNotEndOfAyah = GetPattern("OptionalStopPatternNotEndOfAyah")
            End If
            Return _OptionalStopPatternNotEndOfAyah
        End Get
    End Property
    Public Shared ReadOnly Property CertainNotStopPattern As String
        Get
            If _CertainNotStopPattern Is Nothing Then
                _CertainNotStopPattern = GetPattern("CertainNotStopPattern")
            End If
            Return _CertainNotStopPattern
        End Get
    End Property
    Public Shared ReadOnly Property OptionalNotStopPattern As String
        Get
            If _OptionalNotStopPattern Is Nothing Then
                _OptionalNotStopPattern = GetPattern("OptionalNotStopPattern")
            End If
            Return _OptionalNotStopPattern
        End Get
    End Property
    Public Shared ReadOnly Property TehMarbutaStopRule As String
        Get
            If _TehMarbutaStopRule Is Nothing Then
                _TehMarbutaStopRule = GetPattern("TehMarbutaStopRule")
            End If
            Return _TehMarbutaStopRule
        End Get
    End Property
    Public Shared ReadOnly Property TehMarbutaContinueRule As String
        Get
            If _TehMarbutaContinueRule Is Nothing Then
                _TehMarbutaContinueRule = GetPattern("TehMarbutaContinueRule")
            End If
            Return _TehMarbutaContinueRule
        End Get
    End Property
    Shared _ArabicLongVowels As String()
    Shared _ArabicTanweens As String()
    Shared _ArabicFathaDammaKasra As String()
    Shared _ArabicStopLetters As String()
    Shared _ArabicSpecialGutteral As String()
    Shared _ArabicSpecialLeadingGutteral As String()
    Shared _ArabicPunctuationSymbols As String()
    Shared _ArabicLetters As String()
    Shared _ArabicSunLettersNoLam As String()
    Shared _ArabicSunLetters As String()
    Shared _ArabicMoonLettersNoVowels As String()
    Shared _ArabicMoonLetters As String()
    Shared _RecitationCombiningSymbols As String()
    Shared _RecitationConnectingFollowerSymbols As String()
    Shared _RecitationSymbols As String()
    Shared _ArabicLettersInOrder As String()
    Shared _ArabicSpecialLetters As String()
    Shared _ArabicHamzas As String()
    Shared _ArabicVowels As String()
    Shared _ArabicTajweed As String()
    Shared _ArabicPunctuation As String()
    Shared _NonArabicLetters As String()
    Shared _WhitespaceSymbols As String()
    Shared _PunctuationSymbols As String()
    Shared _RecitationDiacritics As String()
    Shared _RecitationLettersDiacritics As String()
    Shared _RecitationSpecialSymbols As String()
    Shared _ArabicLeadingGutterals As String()
    Shared _RecitationLetters As String()
    Shared _ArabicTrailingGutterals As String()
    Shared _RecitationSpecialSymbolsNotStop As String()
    Public Shared ReadOnly Property ArabicLongVowels As String()
        Get
            If _ArabicLongVowels Is Nothing Then
                _ArabicLongVowels = GetGroup("ArabicLongVowels")
            End If
            Return _ArabicLongVowels
        End Get
    End Property
    Public Shared ReadOnly Property ArabicTanweens As String()
        Get
            If _ArabicTanweens Is Nothing Then
                _ArabicTanweens = GetGroup("ArabicTanweens")
            End If
            Return _ArabicTanweens
        End Get
    End Property
    Public Shared ReadOnly Property ArabicFathaDammaKasra As String()
        Get
            If _ArabicFathaDammaKasra Is Nothing Then
                _ArabicFathaDammaKasra = GetGroup("ArabicFathaDammaKasra")
            End If
            Return _ArabicFathaDammaKasra
        End Get
    End Property
    Public Shared ReadOnly Property ArabicStopLetters As String()
        Get
            If _ArabicStopLetters Is Nothing Then
                _ArabicStopLetters = GetGroup("ArabicStopLetters")
            End If
            Return _ArabicStopLetters
        End Get
    End Property
    Public Shared ReadOnly Property ArabicSpecialGutteral As String()
        Get
            If _ArabicSpecialGutteral Is Nothing Then
                _ArabicSpecialGutteral = GetGroup("ArabicSpecialGutteral")
            End If
            Return _ArabicSpecialGutteral
        End Get
    End Property
    Public Shared ReadOnly Property ArabicSpecialLeadingGutteral As String()
        Get
            If _ArabicSpecialLeadingGutteral Is Nothing Then
                _ArabicSpecialLeadingGutteral = GetGroup("ArabicSpecialLeadingGutteral")
            End If
            Return _ArabicSpecialLeadingGutteral
        End Get
    End Property
    Public Shared ReadOnly Property ArabicPunctuationSymbols As String()
        Get
            If _ArabicPunctuationSymbols Is Nothing Then
                _ArabicPunctuationSymbols = GetGroup("ArabicPunctuationSymbols")
            End If
            Return _ArabicPunctuationSymbols
        End Get
    End Property
    Public Shared ReadOnly Property ArabicLetters As String()
        Get
            If _ArabicLetters Is Nothing Then
                _ArabicLetters = GetGroup("ArabicLetters")
            End If
            Return _ArabicLetters
        End Get
    End Property
    Public Shared ReadOnly Property ArabicSunLettersNoLam As String()
        Get
            If _ArabicSunLettersNoLam Is Nothing Then
                _ArabicSunLettersNoLam = GetGroup("ArabicSunLettersNoLam")
            End If
            Return _ArabicSunLettersNoLam
        End Get
    End Property
    Public Shared ReadOnly Property ArabicSunLetters As String()
        Get
            If _ArabicSunLetters Is Nothing Then
                _ArabicSunLetters = GetGroup("ArabicSunLetters")
            End If
            Return _ArabicSunLetters
        End Get
    End Property
    Public Shared ReadOnly Property ArabicMoonLettersNoVowels As String()
        Get
            If _ArabicMoonLettersNoVowels Is Nothing Then
                _ArabicMoonLettersNoVowels = GetGroup("ArabicMoonLettersNoVowels")
            End If
            Return _ArabicMoonLettersNoVowels
        End Get
    End Property
    Public Shared ReadOnly Property ArabicMoonLetters As String()
        Get
            If _ArabicMoonLetters Is Nothing Then
                _ArabicMoonLetters = GetGroup("ArabicMoonLetters")
            End If
            Return _ArabicMoonLetters
        End Get
    End Property
    Public Shared ReadOnly Property RecitationCombiningSymbols As String()
        Get
            If _RecitationCombiningSymbols Is Nothing Then
                _RecitationCombiningSymbols = GetGroup("RecitationCombiningSymbols")
            End If
            Return _RecitationCombiningSymbols
        End Get
    End Property
    Public Shared ReadOnly Property RecitationConnectingFollowerSymbols As String()
        Get
            If _RecitationConnectingFollowerSymbols Is Nothing Then
                _RecitationConnectingFollowerSymbols = GetGroup("RecitationConnectingFollowerSymbols")
            End If
            Return _RecitationConnectingFollowerSymbols
        End Get
    End Property
    Public Shared ReadOnly Property RecitationSymbols As String()
        Get
            If _RecitationSymbols Is Nothing Then
                _RecitationSymbols = GetGroup("RecitationSymbols")
            End If
            Return _RecitationSymbols
        End Get
    End Property
    Public Shared ReadOnly Property ArabicLettersInOrder As String()
        Get
            If _ArabicLettersInOrder Is Nothing Then
                _ArabicLettersInOrder = GetGroup("ArabicLettersInOrder")
            End If
            Return _ArabicLettersInOrder
        End Get
    End Property
    Public Shared ReadOnly Property ArabicSpecialLetters As String()
        Get
            If _ArabicSpecialLetters Is Nothing Then
                _ArabicSpecialLetters = GetGroup("ArabicSpecialLetters")
            End If
            Return _ArabicSpecialLetters
        End Get
    End Property
    Public Shared ReadOnly Property ArabicHamzas As String()
        Get
            If _ArabicHamzas Is Nothing Then
                _ArabicHamzas = GetGroup("ArabicHamzas")
            End If
            Return _ArabicHamzas
        End Get
    End Property
    Public Shared ReadOnly Property ArabicVowels As String()
        Get
            If _ArabicVowels Is Nothing Then
                _ArabicVowels = GetGroup("ArabicVowels")
            End If
            Return _ArabicVowels
        End Get
    End Property
    Public Shared ReadOnly Property ArabicTajweed As String()
        Get
            If _ArabicTajweed Is Nothing Then
                _ArabicTajweed = GetGroup("ArabicTajweed")
            End If
            Return _ArabicTajweed
        End Get
    End Property
    Public Shared ReadOnly Property ArabicPunctuation As String()
        Get
            If _ArabicPunctuation Is Nothing Then
                _ArabicPunctuation = GetGroup("ArabicPunctuation")
            End If
            Return _ArabicPunctuation
        End Get
    End Property
    Public Shared ReadOnly Property NonArabicLetters As String()
        Get
            If _NonArabicLetters Is Nothing Then
                _NonArabicLetters = GetGroup("NonArabicLetters")
            End If
            Return _NonArabicLetters
        End Get
    End Property
    Public Shared ReadOnly Property WhitespaceSymbols As String()
        Get
            If _WhitespaceSymbols Is Nothing Then
                _WhitespaceSymbols = GetGroup("WhitespaceSymbols")
            End If
            Return _WhitespaceSymbols
        End Get
    End Property
    Public Shared ReadOnly Property PunctuationSymbols As String()
        Get
            If _PunctuationSymbols Is Nothing Then
                _PunctuationSymbols = GetGroup("PunctuationSymbols")
            End If
            Return _PunctuationSymbols
        End Get
    End Property
    Public Shared ReadOnly Property RecitationDiacritics As String()
        Get
            If _RecitationDiacritics Is Nothing Then
                _RecitationDiacritics = GetGroup("RecitationDiacritics")
            End If
            Return _RecitationDiacritics
        End Get
    End Property
    Public Shared ReadOnly Property RecitationLettersDiacritics As String()
        Get
            If _RecitationLettersDiacritics Is Nothing Then
                _RecitationLettersDiacritics = GetGroup("RecitationLettersDiacritics")
            End If
            Return _RecitationLettersDiacritics
        End Get
    End Property
    Public Shared ReadOnly Property RecitationSpecialSymbols As String()
        Get
            If _RecitationSpecialSymbols Is Nothing Then
                _RecitationSpecialSymbols = GetGroup("RecitationSpecialSymbols")
            End If
            Return _RecitationSpecialSymbols
        End Get
    End Property
    Public Shared ReadOnly Property ArabicLeadingGutterals As String()
        Get
            If _ArabicLeadingGutterals Is Nothing Then
                _ArabicLeadingGutterals = GetGroup("ArabicLeadingGutterals")
            End If
            Return _ArabicLeadingGutterals
        End Get
    End Property
    Public Shared ReadOnly Property RecitationLetters As String()
        Get
            If _RecitationLetters Is Nothing Then
                _RecitationLetters = GetGroup("RecitationLetters")
            End If
            Return _RecitationLetters
        End Get
    End Property
    Public Shared ReadOnly Property ArabicTrailingGutterals As String()
        Get
            If _ArabicTrailingGutterals Is Nothing Then
                _ArabicTrailingGutterals = GetGroup("ArabicTrailingGutterals")
            End If
            Return _ArabicTrailingGutterals
        End Get
    End Property
    Public Shared ReadOnly Property RecitationSpecialSymbolsNotStop As String()
        Get
            If _RecitationSpecialSymbolsNotStop Is Nothing Then
                _RecitationSpecialSymbolsNotStop = GetGroup("RecitationSpecialSymbolsNotStop")
            End If
            Return _RecitationSpecialSymbolsNotStop
        End Get
    End Property
    Public Shared Function TranslateRegEx(Value As String, bAll As Boolean) As String
        Return System.Text.RegularExpressions.Regex.Replace(Value, "\{(.*?)\}",
            Function(Match As System.Text.RegularExpressions.Match)
                If bAll Then
                    If Match.Groups(1).Value = "CertainStopPattern" Then Return CertainStopPattern
                    If Match.Groups(1).Value = "OptionalStopPattern" Then Return OptionalStopPattern
                    If Match.Groups(1).Value = "OptionalStopPatternNotEndOfAyah" Then Return OptionalStopPatternNotEndOfAyah
                    If Match.Groups(1).Value = "CertainNotStopPattern" Then Return CertainNotStopPattern
                    If Match.Groups(1).Value = "OptionalNotStopPattern" Then Return OptionalNotStopPattern
                    If Match.Groups(1).Value = "TehMarbutaStopRule" Then Return TehMarbutaStopRule
                    If Match.Groups(1).Value = "TehMarbutaContinueRule" Then Return TehMarbutaContinueRule

                    If Match.Groups(1).Value = "ArabicUniqueLetters" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicUniqueLetters, Function(Str As String) ArabicData.TransliterateFromBuckwalter(Str).Replace(CStr(ArabicData.ArabicMaddahAbove), String.Empty)))
                    If Match.Groups(1).Value = "ArabicNumbers" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicNumbers, Function(Str As String) ArabicData.TransliterateFromBuckwalter(Str)))
                    If Match.Groups(1).Value = "ArabicWaslKasraExceptions" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicWaslKasraExceptions, Function(Str As String) ArabicData.TransliterateFromBuckwalter(Str)))
                    'If Match.Groups(1).Value = "SimpleSuperscriptAlefBefore" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(Arabic.SimpleSuperscriptAlefBefore, Function(Str As String) ArabicData.TransliterateFromBuckwalter(Str.Replace(".", String.Empty).Replace("""", String.Empty).Replace("@", String.Empty).Replace("[", String.Empty).Replace("]", String.Empty).Replace("-", String.Empty).Replace("^", String.Empty))))
                    'If Match.Groups(1).Value = "SimpleSuperscriptAlefNotBefore" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(Arabic.SimpleSuperscriptAlefNotBefore, Function(Str As String) ArabicData.TransliterateFromBuckwalter(Str.Replace(".", String.Empty).Replace("""", String.Empty).Replace("@", String.Empty).Replace("[", String.Empty).Replace("]", String.Empty).Replace("-", String.Empty).Replace("^", String.Empty))))
                    'If Match.Groups(1).Value = "SimpleSuperscriptAlefAfter" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(Arabic.SimpleSuperscriptAlefAfter, Function(Str As String) ArabicData.TransliterateFromBuckwalter(Str.Replace(".", String.Empty).Replace("""", String.Empty).Replace("@", String.Empty).Replace("[", String.Empty).Replace("]", String.Empty).Replace("-", String.Empty).Replace("^", String.Empty))))
                    'If Match.Groups(1).Value = "SimpleSuperscriptAlefNotAfter" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(Arabic.SimpleSuperscriptAlefNotAfter, Function(Str As String) ArabicData.TransliterateFromBuckwalter(Str.Replace(".", String.Empty).Replace("""", String.Empty).Replace("@", String.Empty).Replace("[", String.Empty).Replace("]", String.Empty).Replace("-", String.Empty).Replace("^", String.Empty))))
                    If Match.Groups(1).Value = "ArabicLongShortVowels" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicLongVowels, Function(StrV As String) ArabicData.MakeUniRegEx(StrV(0) + "(?=" + ArabicData.MakeUniRegEx(StrV(1)) + ")")))
                    If Match.Groups(1).Value = "ArabicTanweens" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "ArabicFathaDammaKasra" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "ArabicStopLetters" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "ArabicSpecialGutteral" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicSpecialGutteral, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "ArabicSpecialLeadingGutteral" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "ArabicPunctuationSymbols" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicPunctuationSymbols, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "ArabicLetters" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "ArabicSunLettersNoLam" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicSunLettersNoLam, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "ArabicSunLetters" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "ArabicMoonLettersNoVowels" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicMoonLettersNoVowels, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "ArabicMoonLetters" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicMoonLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "RecitationCombiningSymbols" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(RecitationCombiningSymbols, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "RecitationConnectingFollowerSymbols" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(RecitationConnectingFollowerSymbols, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "PunctuationSymbols" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(PunctuationSymbols, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "RecitationLettersDiacritics" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(RecitationLettersDiacritics, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    If Match.Groups(1).Value = "RecitationSpecialSymbolsNotStop" Then Return ArabicData.MakeRegMultiEx(Array.ConvertAll(RecitationSpecialSymbolsNotStop, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                End If
                If System.Text.RegularExpressions.Regex.Match(Match.Groups(1).Value, "0x([0-9a-fA-F]{4})").Success Then
                    Return If(bAll, ArabicData.MakeUniRegEx(ChrW(Integer.Parse(Match.Groups(1).Value.Substring(2), System.Globalization.NumberStyles.HexNumber))), ChrW(Integer.Parse(Match.Groups(1).Value.Substring(2), System.Globalization.NumberStyles.HexNumber)))
                End If
                For Count = 0 To ArabicData.Data.ArabicLetters.Length - 1
                    If Match.Groups(1).Value = ArabicData.Data.ArabicLetters(Count).UnicodeName Then Return If(bAll, ArabicData.MakeUniRegEx(ArabicData.Data.ArabicLetters(Count).Symbol), ArabicData.Data.ArabicLetters(Count).Symbol)
                Next
                '{0} ignore
                Return Match.Value
            End Function)
    End Function
    Public Shared ReadOnly Property RulesOfRecitationRegEx As IslamData.RuleMetadataTranslation()
        Get
            If _RulesOfRecitationRegEx Is Nothing Then
                _RulesOfRecitationRegEx = CachedData.IslamData.MetaRules
                For SubCount As Integer = 0 To _RulesOfRecitationRegEx.Length - 1
                    _RulesOfRecitationRegEx(SubCount).Match = TranslateRegEx(_RulesOfRecitationRegEx(SubCount).Match, True)
                Next
            End If
            Return _RulesOfRecitationRegEx
        End Get
    End Property
    Public Shared ReadOnly Property WarshScript As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _WarshScript Is Nothing Then
                _WarshScript = GetRuleSet("WarshScript")
            End If
            Return _WarshScript
        End Get
    End Property
    Public Shared ReadOnly Property UthmaniMinimalScript As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _UthmaniMinimalScript Is Nothing Then
                _UthmaniMinimalScript = GetRuleSet("UthmaniMinimalScript")
            End If
            Return _UthmaniMinimalScript
        End Get
    End Property
    Public Shared ReadOnly Property SimpleEnhancedScript As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _SimpleEnhancedScript Is Nothing Then
                _SimpleEnhancedScript = GetRuleSet("SimpleEnhancedScript")
            End If
            Return _SimpleEnhancedScript
        End Get
    End Property
    Public Shared ReadOnly Property SimpleScript As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _SimpleScript Is Nothing Then
                _SimpleScript = GetRuleSet("SimpleScript")
            End If
            Return _SimpleScript
        End Get
    End Property
    Public Shared ReadOnly Property SimpleCleanScript As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _SimpleCleanScript Is Nothing Then
                _SimpleCleanScript = GetRuleSet("SimpleCleanScript")
            End If
            Return _SimpleCleanScript
        End Get
    End Property
    Public Shared ReadOnly Property SimpleMinimalScript As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _SimpleMinimalScript Is Nothing Then
                _SimpleMinimalScript = GetRuleSet("SimpleMinimalScript")
            End If
            Return _SimpleMinimalScript
        End Get
    End Property
    Public Shared ReadOnly Property RomanizationRules As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _RomanizationRules Is Nothing Then
                _RomanizationRules = GetRuleSet("RomanizationRules")
            End If
            Return _RomanizationRules
        End Get
    End Property
    Public Shared ReadOnly Property ColoringSpelledOutRules As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _ColoringSpelledOutRules Is Nothing Then
                _ColoringSpelledOutRules = GetRuleSet("ColoringSpelledOutRules")
            End If
            Return _ColoringSpelledOutRules
        End Get
    End Property
    Public Shared ReadOnly Property ErrorCheckRules As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _ErrorCheck Is Nothing Then
                _ErrorCheck = GetRuleSet("ErrorCheck")
            End If
            Return _ErrorCheck
        End Get
    End Property
	public static $_XMLDocMain = null;
	public static function XMLDocMain() { if ($_XMLDocMain === null) $_XMLDocMain = simplexml_load_file(dirname(__FILE__) . "/metadata/" . TanzilReader::$QuranTextNames[0] . ".xml"); return $_XMLDocMain; }
};

class RenderArray
{
	abstract class RenderTypes {
		const eHeaderLeft = 0;
		const eHeaderCenter = 1;
		const eHeaderRight = 2;
		const eText = 3;
		const eInteractive = 4;
	}
	abstract class RenderDisplayClass {
		const eNested = 0;
		const eArabic = 1;
		const eTransliteration = 2;
		const eLTR = 3;
		const eRTL = 4;
		const eContinueStop = 5;
		const eRanking = 6;
		const eList = 7;
	}
	class RenderText
	{
		public $displayClass;
		public $clr;
		public $text;
		public $font;
		void __construct($_displayClass, $_clr, $_text, $_font)
		{
			$displayClass = $_displayClass;
			$clr = $_clr;
			$text = $_text;
			$font = $_font;
		}
	};
	class RenderItem
	{
		public $type;
		public $textitems;
		void __construct($_type, $_textitems)
		{
			$type = $_type;
			$textitems = $_textitems;
		}
	};
	public static $items = null;
	public static function DoRender($items)
	{
		for ($count = 0; $count < count($items); $count++) {
			for ($index = 0; $index < count($items[count]->textitems); $index++) {
				if ($items[count]->textitems[$index]->displayClass == RenderDisplayClass::eArabic) {
				} elseif ($items[count]->textitems[$index]->displayClass == RenderDisplayClass::eTransliteration) {
				} else {
				}
			}
		}
	}
};

class Arabic
{
	class StringLengthComparer
	{
		private $_scheme;
		void __construct($scheme) { $_scheme = $scheme; }
		public static function Compare($x, $y) { $compare = strlen(GetSchemeValueFromSymbol($x, $_scheme)) - strlen(GetSchemeValueFromSymbol($y, $_scheme)); if ($compare == 0) { $compare = strcmp(GetSchemeValueFromSymbol($x, $_scheme), GetSchemeValueFromSymbol($y, $_scheme)); } return $compare; } }
	}
	public static function TransliterateToScheme($arabicString, $schemeType, $scheme)
	{
		if ($schemeType == ArabicData::TranslitScheme::RuleBased) {
			return Arabic::TransliterateWithRules($arabicString, $scheme, null);
		} elseif ($schemeType == ArabicData::TranslitScheme::Literal) {
			return Arabic::TransliterateToRoman($arabicString, $scheme);
		} else {
			return implode(array_filter(str_split($arabicString), function($check) { return $check == " "; }));
		}
	}
    public static function GetSchemeSpecialValue($index, $scheme)
    {
        $sch = null;
        for ($count = 0; $count < count(CachedData::IslamData::TranslitSchemes); $count++) {
            if (CachedData::IslamData::TranslitSchemes[$count]->Name == $scheme) {
                $sch = CachedData::IslamData::TranslitSchemes[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData::TranslitSchemes) { return ""; }
        return $sch->SpecialLetters[$index];
    }
    public static function GetSchemeSpecialFromMatch($str, $scheme, $bExp)
    {
    	$sch = null;
        for ($count = 0; $count < count(CachedData::IslamData::TranslitSchemes); $count++) {
            if (CachedData::IslamData::TranslitSchemes[$count]->Name == $scheme) {
                $sch = CachedData::IslamData::TranslitSchemes[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData::TranslitSchemes)) { return -1; }
        if ($bExp) {
            for ($count = 0; $count < count(CachedData::ArabicSpecialLetters); $count++) {
                if (preg_match(CachedData::ArabicSpecialLetters[$count], $str) == 1) { return $count; }
            }
        } else {
            if (array_search($str, CachedData::ArabicSpecialLetters) !== false) {
                return array_search($str, CachedData::ArabicSpecialLetters);
            }
        }
        return -1;
    }
    public static function GetSchemeLongVowelFromString($str, $scheme)
    {
        $sch = null;
        for ($count = 0; $count < count(CachedData::IslamData::TranslitSchemes); $count++) {
            if (CachedData::IslamData::TranslitSchemes[$count]->Name == $scheme) {
                $sch = CachedData::IslamData::TranslitSchemes[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData::TranslitSchemes)) { return ""; }
        if (array_search($str, CachedData::ArabicVowels) !== false) {
            return $sch->Vowels(array_search($str, CachedData::ArabicVowels));
        }
        return "";
    }
    public static function GetSchemeGutteralFromString($str, $scheme, $leading)
    {
        $sch = null;
        for ($count = 0; $count < count(CachedData::IslamData::TranslitSchemes); $count++) {
            if (CachedData::IslamData::TranslitSchemes[$count]->Name == $scheme) {
                $sch = CachedData::IslamData::TranslitSchemes[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData::TranslitSchemes)) { return ""; }
        if (array_search($str, CachedData::ArabicLeadingGutterals) !== false) {
            return $sch->Vowels(array_search($str, CachedData::ArabicLeadingGutterals) + count(CachedData::ArabicVowels) + ($leading ? count(CachedData::ArabicLeadingGutterals) : 0));
        }
        return "";
    }
    public static function GetSchemeValueFromSymbol($symbol, $scheme)
    {
        $sch = null;
        for ($count = 0; $count < count(CachedData::IslamData::TranslitSchemes); $count++) {
            if (CachedData::IslamData::TranslitSchemes[$count]->Name == $scheme) {
                $sch = CachedData::IslamData::TranslitSchemes[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData::TranslitSchemes)) { return ""; }
        if (array_search($symbol->Symbol, CachedData::ArabicLettersInOrder) !== false) {
            return $sch->Alphabet(array_search($symbol->Symbol, CachedData::ArabicLettersInOrder));
        } elseif (array_search($symbol->Symbol, CachedData::ArabicHamzas) !== false) {
            return $sch->Hamza(array_search($symbol->Symbol, CachedData::ArabicHamzas));
        } elseif (array_search($symbol->Symbol, CachedData::ArabicVowels) !== false) {
            return $sch->Vowels(array_search($symbol->Symbol, CachedData::ArabicVowels));
        } elseif (array_search($symbol->Symbol, CachedData::ArabicTajweed) !== false) {
            return $sch->Tajweed(array_search($symbol->Symbol, CachedData::ArabicTajweed));
        } elseif (array_search($symbol->Symbol, CachedData::ArabicPunctuation) !== false) {
            return $sch->Punctuation(array_search($symbol->Symbol, CachedData::ArabicPunctuation));
        } elseif (array_search($symbol->Symbol, CachedData::NonArabicLetters) !== false) {
            return $sch->NonArabic(array_search($symbol->Symbol, CachedData::NonArabicLetters));
        }
        return "";
    }
	public static function TransliterateToRoman($arabicString, $scheme)
	{
        $romanString = "";
        $letters = clone(ArabicData::Data::ArabicLetters);
        usort($letters, (new Arabic::StringLengthComparer($scheme))->Compare);
        for ($count = 0; $count < strlen($arabicString); $count++) {
            if ($arabicString[$count] == "\\") {
                $count += 1;
                if ($arabicString[$count] == ",") {
                    $romanString .= ArabicData::ArabicComma;
                } elseif ($arabicString[$count] == ";") {
                    $romanString .= ArabicData::ArabicSemicolon;
                } elseif ($arabicString[$count] == "?") {
                    $romanString .= ArabicData::ArabicQuestionMark;
                } else {
                    $romanString .= $arabicString[$count];
                }
            } else {
                for ($index = 0; $index < count(Letters); $index++) {
                    if ($arabicString[$count] == $letters[Index]->Symbol) {
                        $romanString .= $scheme == "" ? $letters[$index]->ExtendedBuckwalterLetter : Arabic::GetSchemeValueFromSymbol($letters[$index], $scheme);
                        break;
                    }
                }
                if ($index == count($letters) {
                    $romanString .= $arabicString[$count];
                }
            }
        }
        return $romanString;
	}
	class RuleMetadata
	{
        void __construct($newIndex, $newLength, $newType)
        {
            $index = $newIndex;
            $length = $newLength;
            $type = $newType;
        }
        public $index;
        public $length;
        public $type;
        public $children;
    };
	public class RuleFuncs
	{
        const eNone = 0;
        const eUpperCase = 1;
        const eSpellNumber = 2;
        const eSpellLetter = 3;
        const eLookupLetter = 4;
        const eLookupLongVowel = 5;
        const eDivideTanween = 6;
        const eDivideLetterSymbol = 7;
        const eStopOption = 8;
        const eLeadingGutteral = 9;
        const eTrailingGutteral = 10;
    };
    public static $ruleFunctions = [
        function($str, $scheme) { return [strtoupper($str)]; },
        function($str, $scheme) { return [Arabic::TransliterateWithRules(ArabicData::TransliterateFromBuckwalter(Arabic::ArabicWordFromNumber(CInt(Arabic::TransliterateToScheme($str, ArabicData::TranslitScheme::Literal, "")), true, false, false)), $scheme, null)]; },
        function($str, $scheme) { return [Arabic::TransliterateWithRules(ArabicLetterSpelling($str, true), $scheme, null)]; },
        function($str, $scheme) { return [Arabic::GetSchemeValueFromSymbol(ArabicData.Data.ArabicLetters(ArabicData::FindLetterBySymbol($str[0])), $scheme)]; },
        function($str, $scheme) { return [Arabic::GetSchemeLongVowelFromString($str, $scheme)]; },
        function($str, $scheme) { return [CachedData::ArabicFathaDammaKasra[array_search($str, CachedData::ArabicTanweens)], ArabicData::ArabicLetterNoon]; },
        function($str, $scheme) { return ["", ""]; },
        function($str, $scheme) { return [""]; },
        function($str, $scheme) { return [Arabic::GetSchemeGutteralFromString($str.Remove(strlen($str) - 1), $scheme, true) + $str[strlen($str) - 1]]; },
        function($str, $scheme) { return [$str[0] + Arabic::GetSchemeGutteralFromString($str.Remove(0, 1), $scheme, false)]; }
    ];
    public static function ArabicLetterSpelling($input, $quranic)
    {
        Dim Output As String = String.Empty
        For Each Ch As Char In Input
            Dim Index As Integer = ArabicData.FindLetterBySymbol(Ch)
            If Index <> -1 AndAlso IsLetter(Index) Then
                If Output <> String.Empty And Not Quranic Then Output += " "
                Output += If(Quranic, ArabicData.Data.ArabicLetters(Index).SymbolName.Remove(ArabicData.Data.ArabicLetters(Index).SymbolName.Length - 1) + If(ArabicData.Data.ArabicLetters(Index).SymbolName.EndsWith("n"), String.Empty, "o"), ArabicData.Data.ArabicLetters(Index).SymbolName)
            ElseIf Index <> -1 AndAlso ArabicData.Data.ArabicLetters(Index).Symbol = ArabicData.ArabicMaddahAbove Then
                If Not Quranic Then Output += Ch
            End If
        Next
        Return ArabicData.TransliterateFromBuckwalter(Output)
    }
    class RuleMetadataComparer
        public function Compare($x, $y)
        {
            if ($x->Index == $y->Index) {
                return $y->Length.CompareTo($x->Length);
            } else {
                return $y->Index.CompareTo($x->Index);
            }
        }
    };
    public static function ArabicWordForLessThanThousand($number, $useClassic, $useAlefHundred)
    {
        Dim Str As String = String.Empty
        Dim HundStr As String = String.Empty
        If Number >= 100 Then
            HundStr = If(UseAlefHundred, CachedData.ArabicBaseHundredNumbers((Number \ 100) - 1).Insert(2, "A"), CachedData.ArabicBaseHundredNumbers((Number \ 100) - 1))
            If (Number Mod 100) = 0 Then Return HundStr
            Number = Number Mod 100
        End If
        If (Number Mod 10) <> 0 And Number <> 11 And Number <> 12 Then
            Str = CachedData.ArabicBaseNumbers((Number Mod 10) - 1)
        End If
        If Number >= 11 AndAlso Number < 20 Then
            If Number = 11 Or Number = 12 Then
                Str += CachedData.ArabicBaseExtraNumbers(Number - 11)
            Else
                Str = Str.Remove(Str.Length - 1) + "a"
            End If
            Str += " Ea$ara"
        ElseIf (Number = 0 And Str = String.Empty) Or Number = 10 Or Number >= 20 Then
            Str = If(Str = String.Empty, String.Empty, Str + " wa") + CachedData.ArabicBaseTenNumbers(Number \ 10)
        End If
        Return If(UseClassic, If(Str = String.Empty, String.Empty, Str + If(HundStr = String.Empty, String.Empty, " wa")) + HundStr, If(HundStr = String.Empty, String.Empty, HundStr + If(Str = String.Empty, String.Empty, " wa")) + Str)
    }
    public static function ArabicWordFromNumber($number, $useClassic, $useAlefHundred, $useMilliard)
    {
        Dim Str As String = String.Empty
        Dim NextStr As String = String.Empty
        Dim CurBase As Integer = 3
        Dim BaseNums As Long() = {1000, 1000000, 1000000000, 1000000000000}
        Dim Bases As String()() = {CachedData.ArabicBaseThousandNumbers, CachedData.ArabicBaseMillionNumbers, If(UseMilliard, CachedData.ArabicBaseMilliardNumbers, CachedData.ArabicBaseBillionNumbers), CachedData.ArabicBaseTrillionNumbers}
        Do
            If Number >= BaseNums(CurBase) And Number < 2 * BaseNums(CurBase) Then
                NextStr = Bases(CurBase)(0)
            ElseIf Number >= 2 * BaseNums(CurBase) And Number < 3 * BaseNums(CurBase) Then
                NextStr = Bases(CurBase)(1)
            ElseIf Number >= 3 * BaseNums(CurBase) And Number < 10 * BaseNums(CurBase) Then
                NextStr = CachedData.ArabicBaseNumbers(CInt(Number \ BaseNums(CurBase) - 1)).Remove(CachedData.ArabicBaseNumbers(CInt(Number \ BaseNums(CurBase) - 1)).Length - 1) + "u " + Bases(CurBase)(2)
            ElseIf Number >= 10 * BaseNums(CurBase) And Number < 11 * BaseNums(CurBase) Then
                NextStr = CachedData.ArabicBaseTenNumbers(1).Remove(CachedData.ArabicBaseTenNumbers(1).Length - 1) + "u " + Bases(CurBase)(2)
            ElseIf Number >= BaseNums(CurBase) Then
                NextStr = Arabic.ArabicWordForLessThanThousand(CInt((Number \ BaseNums(CurBase)) Mod 100), UseClassic, UseAlefHundred)
                If Number >= 100 * BaseNums(CurBase) And Number < If(UseClassic, 200, 101) * BaseNums(CurBase) Then
                    NextStr = NextStr.Remove(NextStr.Length - 1) + "u " + Bases(CurBase)(0).Remove(Bases(CurBase)(0).Length - 1) + "K"
                ElseIf Number >= 200 * BaseNums(CurBase) And Number < If(UseClassic, 300, 201) * BaseNums(CurBase) Then
                    NextStr = NextStr.Remove(NextStr.Length - 2) + " " + Bases(CurBase)(0).Remove(Bases(CurBase)(0).Length - 1) + "K"
                ElseIf Number >= 300 * BaseNums(CurBase) And (UseClassic Or (Number \ BaseNums(CurBase)) Mod 100 = 0) Then
                    NextStr = NextStr.Remove(NextStr.Length - 1) + "i " + Bases(CurBase)(0).Remove(Bases(CurBase)(0).Length - 1) + "K"
                Else
                    NextStr += " " + Bases(CurBase)(0).Remove(Bases(CurBase)(0).Length - 1) + "FA"
                End If
            End If
            Number = Number Mod BaseNums(CurBase)
            CurBase -= 1
            Str = If(UseClassic, If(NextStr = String.Empty, String.Empty, NextStr + If(Str = String.Empty, String.Empty, " wa")) + Str, If(Str = String.Empty, String.Empty, Str + If(NextStr = String.Empty, String.Empty, " wa")) + NextStr)
            NextStr = String.Empty
        Loop While CurBase >= 0
        If Number <> 0 Or Str = String.Empty Then NextStr = Arabic.ArabicWordForLessThanThousand(CInt(Number), UseClassic, UseAlefHundred)
        Return If(UseClassic, If(NextStr = String.Empty, String.Empty, NextStr + If(Str = String.Empty, String.Empty, " wa")) + Str, If(Str = String.Empty, String.Empty, Str + If(NextStr = String.Empty, String.Empty, " wa")) + NextStr)
    }
    public shared function ReplaceMetadata($arabicString, $metadataRule, $scheme, $optionalStops)
    {
        For Count As Integer = 0 To CachedData.ColoringSpelledOutRules.Length - 1
            Dim Match As String = Array.Find(CachedData.ColoringSpelledOutRules(Count).Match.Split("|"c), Function(Str As String) Array.IndexOf(Array.ConvertAll(MetadataRule.Type.Split("|"c), Function(S As String) System.Text.RegularExpressions.Regex.Replace(S, "\(.*\)", String.Empty)), Str) <> -1)
            If Match <> Nothing Then
                Dim Str As String = String.Format(CachedData.ColoringSpelledOutRules(Count).Evaluator, ArabicString.Substring(MetadataRule.Index, MetadataRule.Length))
                If CachedData.ColoringSpelledOutRules(Count).RuleFunc <> RuleFuncs.eNone Then
                    Dim Args As String() = RuleFunctions(CachedData.ColoringSpelledOutRules(Count).RuleFunc - 1)(Str, Scheme)
                    If Args.Length = 1 Then
                        Str = Args(0)
                    Else
                        Dim MetaArgs As String() = System.Text.RegularExpressions.Regex.Match(MetadataRule.Type, Match + "\((.*)\)").Groups(1).Value.Split(","c)
                        Str = String.Empty
                        For Index As Integer = 0 To Args.Length - 1
                            If Not Args(Index) Is Nothing Then
                                Str += ReplaceMetadata(Args(Index), New RuleMetadata(0, Args(Index).Length, MetaArgs(Index).Replace(" "c, "|"c)), Scheme, OptionalStops)
                            End If
                        Next
                    End If
                End If
                ArabicString = ArabicString.Insert(MetadataRule.Index + MetadataRule.Length, Str).Remove(MetadataRule.Index, MetadataRule.Length)
            End If
        Next
        Return ArabicString
    }
    public static function DoErrorCheck($arabicString)
    {
        //need to check for decomposed first
        Dim Count As Integer
        For Count = 0 To CachedData.ErrorCheckRules.Length - 1
            Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ArabicString, CachedData.ErrorCheckRules(Count).Match)
            For MatchIndex As Integer = 0 To Matches.Count - 1
                If CachedData.ErrorCheckRules(Count).NegativeMatch Is Nothing OrElse Matches(MatchIndex).Result(CachedData.ErrorCheckRules(Count).NegativeMatch) = String.Empty Then
                    //Debug.Print(ErrorCheckRules(Count).Rule + ": " + TransliterateToScheme(ArabicString, ArabicData.TranslitScheme.Literal, String.Empty).Insert(Matches(MatchIndex).Index, "<!-- -->"))
                End If
            Next
        Next
    }
    public shared function JoinContig($arabicString, $preString, $postString)
    {
        Dim Index As Integer = PreString.LastIndexOf(" "c)
        'take last word of pre string and first word of post string or another if it is a pause marker
        'end of ayah sign without number is used as a proper place holder
        If Index <> -1 And PreString.Length - 2 = Index Then Index = PreString.LastIndexOf(" "c, Index - 1)
        If Index <> -1 Then PreString = PreString.Substring(Index + 1)
        If PreString <> String.Empty Then PreString += " " + ArabicData.ArabicEndOfAyah + " "
        Index = PostString.IndexOf(" "c)
        If Index = 2 Then Index = PreString.IndexOf(" "c, Index + 1)
        If Index <> -1 Then PostString = PostString.Substring(0, Index)
        If PostString <> String.Empty Then PostString = " " + ArabicData.ArabicEndOfAyah + " " + PostString
        Return PreString + ArabicString + PostString
    }
    public static function UnjoinContig($arabicString, $preString, $postString)
    {
        $index = $arabicString.IndexOf(ArabicData::ArabicEndOfAyah)
        If PreString <> String.Empty AndAlso Index <> -1 Then
            ArabicString = ArabicString.Substring(Index + 1 + 1)
        End If
        Index = ArabicString.LastIndexOf(ArabicData::ArabicEndOfAyah)
        If PostString <> String.Empty AndAlso Index <> -1 Then
            $arabicString = $arabicString.Substring(0, Index - 1);
        End If
        return $arabicString;
    }
    public static function TransliterateContigWithRules($arabicString, $preString, $postString, $scheme, $optionalStops, $preOptionalStops, $postOptionalStops)
    {
        return Arabic::UnjoinContig(Arabic::TransliterateWithRules(Arabic::JoinContig($arabicString, $preString, $postString), $scheme, null), $preString, $postString);
    }
    public static function TransliterateWithRules($arabicString, $scheme, $optionalStops)
        $metadataList = array();
        DoErrorCheck($arabicString);
        for ($count = 0; $count < count(CachedData::RulesOfRecitationRegEx); $count++) {
            if (CachedData::RulesOfRecitationRegEx[$count]->Evaluator != null) {
                $matches = preg_match_all($arabicString, CachedData::RulesOfRecitationRegEx[$count]->Match);
                for ($matchIndex = 0; $matchIndex < count($matches); $matchIndex) {
                    for ($subCount = 0; $subCount < count(CachedData::RulesOfRecitationRegEx[$count]->Evaluator); $subCount++) {
                        if (CachedData::RulesOfRecitationRegEx[$count]->Evaluator[$subCount] != null && ($matches[$matchIndex]->Groups[$subCount + 1].Length <> 0 Or array_search($allowZeroLength, CachedData::RulesOfRecitationRegEx[$count]->Evaluator[$subCount]) !== false) {
                            $metadataList.push(new RuleMetadata($matches[$matchIndex]->Groups[$subCount + 1].Index, $matches[$matchIndex]->Groups[$subCount + 1].Length, CachedData::RulesOfRecitationRegEx[$count]->Evaluator[$subCount]));
                        }
                    }
                }
            }
        }
        sort(MetadataList, New RuleMetadataComparer);
        for ($index = 0; $index < count(MetadataList); $index++) {
            $arabicString = Arabic::ReplaceMetadata($arabicString, $metadataList[$index], $scheme, $optionalStops);
        }
        //redundant romanization rules should have -'s such as seen/teh/kaf-heh
        for ($count = 0; $count < count(CachedData.RomanizationRules); $count++) {
            if (CachedData::RomanizationRules[$count]->RuleFunc == RuleFuncs::eNone) {
                $arabicString = preg_replace(CachedData::RomanizationRules[$count]->Match, CachedData::RomanizationRules[$count]->Evaluator, $arabicString);
            } else {
                $arabicString = preg_replace_callback(CachedData::RomanizationRules[$count]->Match, function($matches) { return $ruleFunctions[CachedData::RomanizationRules[$count]->RuleFunc - 1](Match.Result(CachedData::RomanizationRules[$count]->Evaluator), $scheme)(0); }, $arabicString);
            }
        }

        //process wasl loanwords and names
        //process loanwords and names
        return $arabicString;
    }
    static function GetTransliterationTable($scheme)
    {
        $items = array();
        $items.AddRange(Array.ConvertAll(CachedData::ArabicLettersInOrder, function($letter) new RenderArray::RenderItem(RenderArray::RenderTypes::eText, New RenderArray::RenderText() {New RenderArray::RenderText(RenderArray::RenderDisplayClass::eArabic, Letter), New RenderArray::RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArabicData.Data.ArabicLetters(ArabicData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        $items.AddRange(Array.ConvertAll(CachedData::ArabicSpecialLetters, function($combo) new RenderArray::RenderItem(RenderArray::RenderTypes::eText, New RenderArray::RenderText() {New RenderArray::RenderText(RenderArray::RenderDisplayClass::eArabic, System.Text.RegularExpressions.Regex.Replace(Combo.Replace(CachedData::TehMarbutaStopRule, String.Empty).Replace(CachedData.TehMarbutaContinueRule, "..."), "\(?\\u([0-9a-fA-F]{4})\)?", Function(Match As System.Text.RegularExpressions.Match) ChrW(Integer.Parse(Match.Groups(1).Value, Globalization.NumberStyles.HexNumber)))), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeSpecialValue(GetSchemeSpecialFromMatch(Combo, Scheme, False), Scheme))})))
        $items.AddRange(Array.ConvertAll(CachedData::ArabicHamzas, function($letter) new RenderArray::RenderItem(RenderArray::RenderTypes::eText, New RenderArray::RenderText() {New RenderArray::RenderText(RenderArray::RenderDisplayClass::eArabic, Letter), New RenderArray::RenderText(RenderArray::RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArabicData.Data.ArabicLetters(ArabicData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        $items.AddRange(Array.ConvertAll(CachedData::ArabicVowels, function($combo) new RenderArray::RenderItem(RenderArray::RenderTypes::eText, New RenderArray::RenderText() {New RenderArray::RenderText(RenderArray::RenderDisplayClass::eArabic, Combo), New RenderArray::RenderText(RenderArray::RenderDisplayClass.eTransliteration, GetSchemeLongVowelFromString(Combo, Scheme))})))
        $items.AddRange(Array.ConvertAll(CachedData::ArabicTajweed, function($letter) new RenderArray::RenderItem(RenderArray::RenderTypes::eText, New RenderArray::RenderText() {New RenderArray::RenderText(RenderArray::RenderDisplayClass::eArabic, Letter), New RenderArray::RenderText(RenderArray::RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArabicData.Data.ArabicLetters(ArabicData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        return $items;
    }
    public static function GetTransliterationSchemes()
    {
        $strings = array();
        $strings[0] = [Utility::LoadResourceString("IslamSource_Off"), "0"];
        $strings[1] = [Utility::LoadResourceString("IslamSource_ExtendedBuckwalter"), "1"];
        for ($count = 0; $count < count(CachedData::IslamData::TranslitSchemes); $count++) {
            $strings[$count * 2 + 2] = [Utility::LoadResourceString("IslamSource_" . CachedData::IslamData::TranslitSchemes[$count]->Name), CStr($count * 2 + 2)];
            $strings[$count * 2 + 1 + 2] = [Utility::LoadResourceString("IslamSource_" . CachedData::IslamData::TranslitSchemes[$count]->Name) . " Literal", CStr($count * 2 + 1 + 2)];
        }
        return $strings;
    }
};
class TanzilReader
{
	public static function IsQuranTextReference($str) { return preg_match("/^(?:,?(\d+)(?:\:(\d+))?(?:\:(\d+))?(?:-(\d+)(?:\:(\d+))?(?:\:(\d+))?)?)+$/s", $str); }
	public static function QuranTextFromReference($str)
	{
		return preg_replace_callback('/(?:,?(\d+)(?:\:(\d+))?(?:\:(\d+))?(?:-(\d+)(?:\:(\d+))?(?:\:(\d+))?)?)/s', function ($matches) {
			$basechapter = $matches[1];
			$baseverse = $matches[2];
			$wordnumber = $matches[3];
			$endchapter = $matches[4];
			$extraversenumber = $matches[5];
			$endwordnumber = $matches[6];
			if ($baseverse != 0 && $wordnumber == 0 && $endchapter != 0 && $extraversenumber == 0 && $endwordnumber == 0) {
	            $extraversenumber = $endchapter;
	            $endchapter = 0;
	        } elseif ($baseverse != 0 && $wordnumber != 0 && $endchapter != 0 && $extraversenumber == 0 && $endwordnumber == 0) {
	            $endwordnumber = $endchapter;
	            $endchapter = 0;
	        } elseif ($baseverse != 0 && $wordnumber != 0 && $endchapter != 0 && $extraversenumber != 0 && $endwordnumber == 0) {
	            $endwordnumber = $extraversenumber;
	            $extraversenumber = $endchapter;
	            $endchapter = 0;
	        }
	        if ($baseverse == 0) {
	            $baseverse += 1;
	            $extraversenumber = count(TanzilReader::GetTextChapter(CachedData::XMLDocMain(), BaseChapter)->children());
	        }
	        if ($wordnumber == 0) $wordnumber += 1;
	        return TanzilReader::DoGetRenderedQuranText(TanzilReader::QuranTextRangeLookup($basechapter, $baseverse, $wordnumber, $endchapter, $extraversenumber, $endwordnumber), $basechapter, $baseverse, null, null, null, null);
		}, $str);
	}
	public static function QuranTextRangeLookup($basechapter, $baseverse, $wordnumber, $endchapter, $extraversenumber, $endwordnumber)
	{
		$qurantext = array();
		if ($endchapter == 0 Or $endchapter == $basechapter) {
			array_push($qurantext, TanzilReader::GetQuranText(CachedData::XMLDocMain(), $basechapter, $baseverse, $extraversenumber != 0 ? $extraversenumber : $baseverse));
		} else {
			array_push($qurantext, TanzilReader::GetQuranTextRange(CachedData::XMLDocMain(), $basechapter, $baseverse, $endchapter, $extraversenumber));
		}
		if ($wordnumber > 1) {
		}
		if ($endwordnumber != 0) {
		}
		return $qurantext;
	}
	public static function GetQuranTextRange($xmldoc, $startchapter, $startverse, $endchapter, $endverse)
	{
		if ($startchapter == -1) $startchapter = 1;
		if ($endchapter == -1) $endchapter = TanzilReader::GetChapterCount();
		$chapterverses = array();
        for ($count = $startchapter; $count <= $endchapter; $count++) {
            $chapterverses.push(TanzilReader::GetQuranText($xmldoc, $count, $startchapter == $count ? $startverse : -1, $endchapter == $count ? $endverse : -1));
        }
        return $chapterverses;
	}
	public static function GetQuranText($xmldoc, $chapter, $startverse, $endverse)
	{
		if ($startverse == -1) $startverse = 1;
		if ($endverse == -1) $endverse = count(TanzilReader::GetTextChapter($xmldoc, $chapter)->children());
		$verses = array();
		for ($count = $startverse; $count <= $endverse; $count++) {
			$versenode = TanzilReader::GetTextVerse(TanzilReader::GetTextChapter($xmldoc, $chapter), $count);
			if ($versenode !== null) {
				if (isset($versenode->attributes()["text"])) {
					array_push($verses, (string)$versenode->attributes()["text"]);
				}
			}
		}
		return $verses;
	}
	public static function GetTranslationIndex($translation)
	{
		if (!$translation) $translation = CachedData::IslamData()->translations->attributes()["default"];
		for ($count = 0; $count < count(CachedData::IslamData()->translations->children()); $count++) {
			if (CachedData::IslamData()->translations->children()[$count]->attributes()["file"] == $translation) return $count;
		}
		$translation = CachedData::IslamData()->translations->attributes()["default"];
		for ($count = 0; $count < count(CachedData::IslamData()->translations->children()); $count++) {
			if (CachedData::IslamData()->translations->children()[$count]->attributes()["file"] == $translation) return $count;
		}
		return -1;
	}
	public static function GetTranslationFileName($translation)
	{
		$index = TanzilReader::GetTranslationIndex($translation);
		return CachedData::IslamData()->translations->children()[$index]->attributes()["file"] + ".txt"
	}
	public static function DoGetRenderedQuranText($qurantext, $basechapter, $baseverse, $translation, $schemetype, $scheme, $translationindex)
	{
		$text = "";
		$lines = explode("\n", file_get_contents(dirname(__FILE__) . "/metadata/" . TanzilReader::GetTranslationFileName($translation)));
		$w4wlines = explode("\n", file_get_contents(dirname(__FILE__) . "/metadata/en.w4w.shehnazshaikh.txt"));
		if ($qurantext !== null) {
			for ($chapter = 0; $chapter < count($qurantext); $chapter++) {
				for ($verse = 0; $verse < count($qurantext[$chapter]); $verse++) {
					$text .= $qurantext[$chapter][$verse];
					//TanzilReader::GetTranslationVerse($lines, $basechapter + $chapter, ($chapter == 0 ? $baseverse : 1) + $verse);
				}
			}
		}
		return $text;
	}
	public static function GetTranslationVerse($lines, $chapter, $verse)
	{
		return $lines[TanzilReader::GetVerseNumber($chapter, $verse) - 1];
	}
	public static function GetVerseNumber($chapter, $verse)
	{
		return TanzilReader::GetChapterByIndex($chapter)->attributes()["start"] + $verse;
	}
	public static function GetChapterByIndex($index)
	{
		return Utility::GetChildNodeByIndex("sura", "index", $index, CachedData.XMLDocInfo()->children()["suras"]->children());
	}
	public static function GetTextChapter($xmldoc, $chapter)
	{
		return Utility::GetChildNodeByIndex("sura", "index", $chapter, $xmldoc->children());
	}
	public static function GetTextVerse($chapternode, $verse)
	{
		return Utility::GetChildNodeByIndex("aya", "index", $verse, $chapternode->children());
	}
	public static function GetChapterCount()
	{
        return Utility::GetChildNodeCount("sura", CachedData.XMLDocInfo()->children()["suras"]);
    }
    public static $QuranTextNames = array("quran-hafs", "quran-warsh", "quran-alduri");
};

class DocBuilder
{
	public static function BuckwalterTextFromReferences($str)
	{
		$matches = array();
		$out = "";
		if (preg_match_all('/(.*?)(?:(\\\{)(.*?)(\\\})|$)/s', CachedData::IslamData()->lists->category->children()[$count]["text"], $matches, PREG_SET_ORDER) !== 0) {
			$val = "";
			for ($matchcount = 0; $matchcount < count($matches); $matchcount++) {
				if (count($matches[$matchcount]) >= 4) {
					$check = DocBuilder::TextFromReferences($matches[$matchcount][3]);
					if ($check !== "") $val .= (($val !== "") ? "," : "") . $check;
				}
			}
			if ($val !== "") $out .= (($out !== "") ? "," : "") . "{\"text\": \"" . Utility::LoadResourceString("IslamInfo_" . (string)CachedData::IslamData()->lists->category->children()[$count]["id"]) . "\", \"menu\":[" . $val . "]}";
		}
		return $out;
	}
	public static function TextFromReferences($str)
	{
		$val = "";
		for ($count = 0; $count < count(CachedData::IslamData()->abbreviations->children()); $count++) {
			for ($subcount = 0; $subcount < count(CachedData::IslamData()->abbreviations->children()[$count]->children()); $subcount++) {
				if (array_search($str, explode("|", (string)CachedData::IslamData()->abbreviations->children()[$count]->children()[$subcount]->attributes()["text"])) !== false) {
					$val = "";
					if ((string)CachedData::IslamData()->abbreviations->children()[$count]->children()[$subcount]->attributes()["font"] != "") {
						foreach (explode("|", (string)CachedData::IslamData()->abbreviations->children()[$count]->children()[$subcount]->attributes()["font"]) as $part) {
							$font = "";
							if (strpos($part, ';') !== false) {
								$font = explode(';', $part)[0];
								$part = explode(';', $part)[1];
							}
							foreach (explode(",", $part) as $substr) {
								$val .= "{\"value\":\"" . $str . "\", \"font\": \"" . $font . "\", \"char\": \"" . implode(",", explode("+", $substr)) . "\"},";
							}
						}
					}
					if ($val !== "") {
						$val = "{\"text\": \"" . Utility::LoadResourceString("IslamInfo_" . (string)CachedData::IslamData()->abbreviations->children()[$count]->children()[$subcount]->attributes()["id"]) . "\", \"values\":[" . substr($val, 0, -1) . "]}";
					}
					break;
				}
			}
		}
		return $val;
	}
};
function getcacheitem()
{
	// Settings
	$cachedir = dirname(__FILE__) . '/cache/'; // Directory to cache files in (keep outside web root)
	$cachetime = 600; // Seconds to cache files for
	$cacheext = 'cache'; // Extension to give cached files (usually 
	$page = 'http://' . $_SERVER['HTTP_HOST'] . $_SERVER['REQUEST_URI']; // Requested page
	$cachefile = $cachedir . md5($page) . '.' . $cacheext; // Cache file to either load or create
	$cachefile_created = (@file_exists($cachefile)) ? @filemtime($cachefile) : 0;
	@clearstatcache();
	if (time() - $cachetime < $cachefile_created) {
		@readfile($cachefile);
		return true;
	}
	return false;
}
function cacheitemimg($img)
{
	$cachedir = dirname(__FILE__) . '/cache/'; // Directory to cache files in (keep outside web root)
	$cacheext = 'cache'; // Extension to give cached files (usually 
	$page = 'http://' . $_SERVER['HTTP_HOST'] . $_SERVER['REQUEST_URI']; // Requested page
	$cachefile = $cachedir . md5($page) . '.' . $cacheext; // Cache file to either load or create
	imagepng($img, $cachefile);
}
function cacheitem($item)
{
	$cachedir = dirname(__FILE__) . '/cache/'; // Directory to cache files in (keep outside web root)
	$cacheext = 'cache'; // Extension to give cached files (usually 
	$page = 'http://' . $_SERVER['HTTP_HOST'] . $_SERVER['REQUEST_URI']; // Requested page
	$cachefile = $cachedir . md5($page) . '.' . $cacheext; // Cache file to either load or create
	$fp = @fopen($cachefile, 'w'); 
    // save the contents of output buffer to the file
    @fwrite($fp, $item);
    @fclose($fp);
}
/*
IslamSource image render tool
*/
function fontFolder() {
    $osName = php_uname( 's' );
    if (strtoupper(substr($osName, 0, 3)) === 'WIN') {
        return '/Windows/Fonts/';
    }
    if (strtoupper(substr($osName, 0, 5)) === 'LINUX') {
        return '/usr/share/fonts/truetype/';
    }
    if (strtoupper(substr($osName, 0, 7)) === 'FREEBSD') {
        //This is not tested
        return '/usr/share/fonts/truetype/';
    }
}
if (array_key_exists("Char", $_GET)) {
	if (!getcacheitem()) {
		//GD library and FreeType library
		$size = $_GET["Size"];
		if ($size === null) $size = 32;
		$chr = mb_convert_encoding("&#" . hexdec($_GET["Char"]) . ";", 'UTF-8', 'HTML-ENTITIES');
		$fonts = array("AGAIslamicPhrases", "AGAArabesque", "Shia", "IslamicLogo", "KFGQPCArabicSymbols01", "Quranic", "Tulth", "Farsi", "Asmaul-Husna", "Asmaul-Husna_2");
		$fontfiles = array("AGA_Islamic_Phrases.TTF", "aga-arabesque.ttf", "SHIA.TTF", "islamic.ttf", "Symbols1_Ver02.otf", "Quranic.ttf", "Tulth.ttf", "Farsi.ttf", "Asmaul-Husna_1.ttf", "Asmaul-Husna_2.ttf");
		$font = dirname(__FILE__) . '/files/' . (array_key_exists("Font", $_GET) && array_search($_GET["Font"], $fonts) !== false ? $fontfiles[array_search($_GET["Font"], $fonts)] : 'me_quran.ttf');
		$bbox = imageftbbox($size, 0, $font, $chr);
		$im = imagecreatetruecolor(ceil($bbox[4] - $bbox[0]), ceil($bbox[1] - $bbox[5]));
		$white = imagecolorallocate($im, 255, 255, 255);
		$black = imagecolorallocate($im, 0, 0, 0);
		imagefilledrectangle($im, 0, 0, ceil($bbox[4] - $bbox[0]), ceil($bbox[1] - $bbox[5]), $white);
		imagefttext($im, $size, 0, $bbox[0], -$bbox[5], $black, $font, $chr);
		header('Content-Type: image/png');
		imagepng($im);
		cacheitemimg($im);
		imagedestroy($im);
	}
} elseif (array_key_exists("Cat", $_GET)) {
	if (!getcacheitem()) {
		$json = "";
		for ($count = 0; $count < count(CachedData::IslamData()->lists->category->children()); $count++) {
			$json .= (($json !== "") ? "," : "") . DocBuilder::BuckwalterTextFromReference(CachedData::IslamData()->lists->category->children()[$count]["text"]);
		}
		$json = "[" . $json . "]";
		echo $json;
		cacheitem($json);
	}
} else {
/* OPTIONS PAGE */

// Add admin actions
add_action('admin_init', 'islamic_source_init');

// Register settings
function islamic_source_init() {
   if ( ! current_user_can('edit_posts') && ! current_user_can('edit_pages') ) {
     return;
   }
 
   if ( get_user_option('rich_editing') == 'true' ) {
     add_filter( 'mce_external_plugins', 'is_add_plugin' );
     
	 $rowvalue = '';
	 add_filter( 'mce_buttons'.$rowvalue, 'is_register_button' );
   }
}
 
 /**
Register Button
*/

function is_register_button( $buttons ) {
	array_push( $buttons, "is_button");
	return $buttons;
}
 
/**
Register IS Plugin
*/
 
function is_add_plugin( $plugin_array ) {
 	$url = plugins_url().""; 
   	$plugin_array['is_button'] = $url.'/IslamSource/isbutton.js';
  	return $plugin_array;
}

/* REPLACE SHORTCODES */

function is_shortcode() {
    global $wp_query;	
    $posts = $wp_query->posts;
    foreach ($posts as $post){
		$post->post_content = preg_replace_callback('/(.*?)(?:(\{)(.*?)(\})|$)/s', function ($matches) {
			if (TanzilReader::IsQuranTextReference($matches[3])) {
				return $matches[1] . "<div style=\"direction: rtl\"><span dir=\"rtl\" style=\"font-size: 40px;\">&#xFD3F;" . TanzilReader::QuranTextFromReference($matches[3]) . "&#xFD3E;</span>&nbsp;(" . $matches[3] . ")</div>";
			} else if (strpos($matches[3], ";Char=") !== false) {
				return "<img src=\"" . plugins_url()."" . "/IslamSource/IslamSourceWP.php?Size=100&" . str_replace(';', '&', substr($matches[3], strpos($matches[3], ';') + 1)) . "\">";
			} else {
				return $matches[0];
			}
		}, $post->post_content);
    }
}
add_action( 'wp', 'is_shortcode' );

}
?>