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

abstract class TranslitScheme
{
    const None = 0;
    const Literal = 1;
    const RuleBased = 2;
};

class Arabic
{
    static $_ArabicXMLData = null;
    public static function Data()
    {
        if (Arabic::$_ArabicXMLData == null) {
            Arabic::$_ArabicXMLData = simplexml_load_file(dirname(__FILE__) . "/HostPage/metadata/arabicdata.xml");
        }
        return Arabic::$_ArabicXMLData;
    }
	public static $_BuckwalterMap = null;
    public static function BuckwalterMap()
	{
        if (Arabic::$_BuckwalterMap == null) {
            Arabic::$_BuckwalterMap = array();
            for ($index = 0; $index < count(Arabic::Data()->ArabicLetters); $index++) {
                if (Arabic::Data()->ArabicLetters[$index]->ExtendedBuckwalterLetter != 0) {
                    Arabic::$_BuckwalterMap[Arabic::Data()->ArabicLetters[$index]->ExtendedBuckwalterLetter] = $index;
                }
            }
        }
        return Arabic::$_BuckwalterMap;
    }
    public static function TransliterateFromBuckwalter($buckwalter)
    {
    	$arabicString = "";
        for ($count = 0; $count < strlen($buckwalter); $count++) {
            if ($buckwalter[$count] == "\\") {
                $count += 1;
                if ($buckwalter[$count] == ",") {
                    $arabicString += ArabicData::ArabicComma;
                } else {
                    $arabicString += $buckwalter[$count];
                }
            } else {
                if (array_key_exists($buckwalter[$count], Arabic::BuckwalterMap())) {
                    $arabicString += Arabic::Data()->ArabicLetters(Arabic::BuckwalterMap()[$buckwalter[$count]])->Symbol;
                } else {
                    $arabicString += $buckwalter[$count];
                }
            }
        }
        return $arabicString;
    }
    public static $_ArabicLetterMap = null;
    public static function ArabicLetterMap()
    {
        if (Arabic::$_ArabicLetterMap == null) {
            Arabic::$_ArabicLetterMap = array();
            for ($index = 0; $index < count(Arabic::Data()->ArabicLetters); $index++) {
                if (Arabic::Data()->ArabicLetters[$index]->Symbol != 0) {
                    Arabic::$_ArabicLetterMap[Arabic::Data()->ArabicLetters[$index]->Symbol] = $index;
                }
            }
        }
        return Arabic::$_ArabicLetterMap;
	}
    public static function FindLetterBySymbol($symbol)
    {
        return array_key_exists($symbol, Arabic::ArabicLetterMap()) ? Arabic::ArabicLetterMap()[$symbol] : -1;
    }
    public static function init() {
        Arabic::$Space = mb_convert_encoding('&#x0020;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ExclamationMark = mb_convert_encoding('&#x0021;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$QuotationMark = mb_convert_encoding('&#x0022;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$Comma = mb_convert_encoding('&#x002C;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$FullStop = mb_convert_encoding('&#x002E;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$HyphenMinus = mb_convert_encoding('&#x002D;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$Colon = mb_convert_encoding('&#x003A;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$LeftParenthesis = mb_convert_encoding('&#x005B;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$RightParenthesis = mb_convert_encoding('&#x005D;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$LeftSquareBracket = mb_convert_encoding('&#x005B;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$RightSquareBracket = mb_convert_encoding('&#x005D;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$LeftCurlyBracket = mb_convert_encoding('&#x007B;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$RightCurlyBracket = mb_convert_encoding('&#x007D;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$NoBreakSpace = mb_convert_encoding('&#x00A0;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$LeftPointingDoubleAngleQuotationMark = mb_convert_encoding('&#x00AB;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$RightPointingDoubleAngleQuotationMark = mb_convert_encoding('&#x00BB;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicComma = mb_convert_encoding('&#x060C;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterHamza = mb_convert_encoding('&#x0621;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterAlefWithMaddaAbove = mb_convert_encoding('&#x0622;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterAlefWithHamzaAbove = mb_convert_encoding('&#x0623;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterWawWithHamzaAbove = mb_convert_encoding('&#x0624;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterAlefWithHamzaBelow = mb_convert_encoding('&#x0625;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterYehWithHamzaAbove = mb_convert_encoding('&#x0626;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterAlef = mb_convert_encoding('&#x0627;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterBeh = mb_convert_encoding('&#x0628;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterTehMarbuta = mb_convert_encoding('&#x0629;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterTeh = mb_convert_encoding('&#x062A;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterTheh = mb_convert_encoding('&#x062B;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterJeem = mb_convert_encoding('&#x062C;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterHah = mb_convert_encoding('&#x062D;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterKhah = mb_convert_encoding('&#x062E;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterDal = mb_convert_encoding('&#x062F;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterThal = mb_convert_encoding('&#x0630;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterReh = mb_convert_encoding('&#x0631;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterZain = mb_convert_encoding('&#x0632;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterSeen = mb_convert_encoding('&#x0633;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterSheen = mb_convert_encoding('&#x0634;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterSad = mb_convert_encoding('&#x0635;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterDad = mb_convert_encoding('&#x0636;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterTah = mb_convert_encoding('&#x0637;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterZah = mb_convert_encoding('&#x0638;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterAin = mb_convert_encoding('&#x0639;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterGhain = mb_convert_encoding('&#x063A;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicTatweel = mb_convert_encoding('&#x0640;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterFeh = mb_convert_encoding('&#x0641;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterQaf = mb_convert_encoding('&#x0642;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterKaf = mb_convert_encoding('&#x0643;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterLam = mb_convert_encoding('&#x0644;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterMeem = mb_convert_encoding('&#x0645;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterNoon = mb_convert_encoding('&#x0646;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterHeh = mb_convert_encoding('&#x0647;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterWaw = mb_convert_encoding('&#x0648;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterAlefMaksura = mb_convert_encoding('&#x0649;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterYeh = mb_convert_encoding('&#x064A;', 'UTF-8', 'HTML-ENTITIES');

        Arabic::$ArabicFathatan = mb_convert_encoding('&#x064B;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicDammatan = mb_convert_encoding('&#x064C;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicKasratan = mb_convert_encoding('&#x064D;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicFatha = mb_convert_encoding('&#x064E;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicDamma = mb_convert_encoding('&#x064F;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicKasra = mb_convert_encoding('&#x0650;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicShadda = mb_convert_encoding('&#x0651;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSukun = mb_convert_encoding('&#x0652;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicMaddahAbove = mb_convert_encoding('&#x0653;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicHamzaAbove = mb_convert_encoding('&#x0654;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicHamzaBelow = mb_convert_encoding('&#x0655;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicVowelSignDotBelow = mb_convert_encoding('&#x065C;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$Bullet = mb_convert_encoding('&#x2022;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterSuperscriptAlef = mb_convert_encoding('&#x0670;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterAlefWasla = mb_convert_encoding('&#x0671;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallHighLigatureSadWithLamWithAlefMaksura = mb_convert_encoding('&#x06D6;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallHighLigatureQafWithLamWithAlefMaksura = mb_convert_encoding('&#x06D7;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallHighMeemInitialForm = mb_convert_encoding('&#x06D8;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallHighLamAlef = mb_convert_encoding('&#x06D9;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallHighJeem = mb_convert_encoding('&#x06DA;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallHighThreeDots = mb_convert_encoding('&#x06DB;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallHighSeen = mb_convert_encoding('&#x06DC;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicEndOfAyah = mb_convert_encoding('&#x06DD;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicStartOfRubElHizb = mb_convert_encoding('&#x06DE;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallHighRoundedZero = mb_convert_encoding('&#x06DF;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallHighUprightRectangularZero = mb_convert_encoding('&#x06E0;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallHighMeemIsolatedForm = mb_convert_encoding('&#x06E2;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallLowSeen = mb_convert_encoding('&#x06E3;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallWaw = mb_convert_encoding('&#x06E5;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallYeh = mb_convert_encoding('&#x06E6;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallHighNoon = mb_convert_encoding('&#x06E8;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicPlaceOfSajdah = mb_convert_encoding('&#x06E9;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicEmptyCentreLowStop = mb_convert_encoding('&#x06EA;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicEmptyCentreHighStop = mb_convert_encoding('&#x06EB;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicRoundedHighStopWithFilledCentre = mb_convert_encoding('&#x06EC;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSmallLowMeem = mb_convert_encoding('&#x06ED;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicSemicolon = mb_convert_encoding('&#x061B;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterMark = mb_convert_encoding('&#x061C;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicQuestionMark = mb_convert_encoding('&#x061F;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterPeh = mb_convert_encoding('&#x067E;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterTcheh = mb_convert_encoding('&#x0686;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterVeh = mb_convert_encoding('&#x06A4;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterGaf = mb_convert_encoding('&#x06AF;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ArabicLetterNoonGhunna = mb_convert_encoding('&#x06BA;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ZeroWidthSpace = mb_convert_encoding('&#x200B;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ZeroWidthNonJoiner = mb_convert_encoding('&#x200C;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$ZeroWidthJoiner = mb_convert_encoding('&#x200D;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$LeftToRightMark = mb_convert_encoding('&#x200E;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$RightToLeftMark = mb_convert_encoding('&#x200F;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$PopDirectionalFormatting = mb_convert_encoding('&#x202C;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$LeftToRightOverride = mb_convert_encoding('&#x202D;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$RightToLeftOverride = mb_convert_encoding('&#x202E;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$NarrowNoBreakSpace = mb_convert_encoding('&#x202F;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$DottedCircle = mb_convert_encoding('&#x25CC;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$OrnateLeftParenthesis = mb_convert_encoding('&#xFD3E;', 'UTF-8', 'HTML-ENTITIES');
        Arabic::$OrnateRightParenthesis = mb_convert_encoding('&#xFD3F;', 'UTF-8', 'HTML-ENTITIES');
    }
    public static $Space;
    public static $ExclamationMark;
    public static $QuotationMark;
    public static $Comma;
    public static $FullStop;
    public static $HyphenMinus;
    public static $Colon;
    public static $LeftParenthesis;
    public static $RightParenthesis;
    public static $LeftSquareBracket;
    public static $RightSquareBracket;
    public static $LeftCurlyBracket;
    public static $RightCurlyBracket;
    public static $NoBreakSpace;
    public static $LeftPointingDoubleAngleQuotationMark;
    public static $RightPointingDoubleAngleQuotationMark;
    public static $ArabicComma;
    public static $ArabicLetterHamza;
    public static $ArabicLetterAlefWithMaddaAbove;
    public static $ArabicLetterAlefWithHamzaAbove;
    public static $ArabicLetterWawWithHamzaAbove;
    public static $ArabicLetterAlefWithHamzaBelow;
    public static $ArabicLetterYehWithHamzaAbove;
    public static $ArabicLetterAlef;
    public static $ArabicLetterBeh;
    public static $ArabicLetterTehMarbuta;
    public static $ArabicLetterTeh;
    public static $ArabicLetterTheh;
    public static $ArabicLetterJeem;
    public static $ArabicLetterHah;
    public static $ArabicLetterKhah;
    public static $ArabicLetterDal;
    public static $ArabicLetterThal;
    public static $ArabicLetterReh;
    public static $ArabicLetterZain;
    public static $ArabicLetterSeen;
    public static $ArabicLetterSheen;
    public static $ArabicLetterSad;
    public static $ArabicLetterDad;
    public static $ArabicLetterTah;
    public static $ArabicLetterZah;
    public static $ArabicLetterAin;
    public static $ArabicLetterGhain;
    public static $ArabicTatweel;
    public static $ArabicLetterFeh;
    public static $ArabicLetterQaf;
    public static $ArabicLetterKaf;
    public static $ArabicLetterLam;
    public static $ArabicLetterMeem;
    public static $ArabicLetterNoon;
    public static $ArabicLetterHeh;
    public static $ArabicLetterWaw;
    public static $ArabicLetterAlefMaksura;
    public static $ArabicLetterYeh;

    public static $ArabicFathatan;
    public static $ArabicDammatan;
    public static $ArabicKasratan;
    public static $ArabicFatha;
    public static $ArabicDamma;
    public static $ArabicKasra;
    public static $ArabicShadda;
    public static $ArabicSukun;
    public static $ArabicMaddahAbove;
    public static $ArabicHamzaAbove;
    public static $ArabicHamzaBelow;
    public static $ArabicVowelSignDotBelow;
    public static $Bullet;
    public static $ArabicLetterSuperscriptAlef;
    public static $ArabicLetterAlefWasla;
    public static $ArabicSmallHighLigatureSadWithLamWithAlefMaksura;
    public static $ArabicSmallHighLigatureQafWithLamWithAlefMaksura;
    public static $ArabicSmallHighMeemInitialForm;
    public static $ArabicSmallHighLamAlef;
    public static $ArabicSmallHighJeem;
    public static $ArabicSmallHighThreeDots;
    public static $ArabicSmallHighSeen;
    public static $ArabicEndOfAyah;
    public static $ArabicStartOfRubElHizb;
    public static $ArabicSmallHighRoundedZero;
    public static $ArabicSmallHighUprightRectangularZero;
    public static $ArabicSmallHighMeemIsolatedForm;
    public static $ArabicSmallLowSeen;
    public static $ArabicSmallWaw;
    public static $ArabicSmallYeh;
    public static $ArabicSmallHighNoon;
    public static $ArabicPlaceOfSajdah;
    public static $ArabicEmptyCentreLowStop;
    public static $ArabicEmptyCentreHighStop;
    public static $ArabicRoundedHighStopWithFilledCentre;
    public static $ArabicSmallLowMeem;
    public static $ArabicSemicolon;
    public static $ArabicLetterMark;
    public static $ArabicQuestionMark;
    public static $ArabicLetterPeh;
    public static $ArabicLetterTcheh;
    public static $ArabicLetterVeh;
    public static $ArabicLetterGaf;
    public static $ArabicLetterNoonGhunna;
    public static $ZeroWidthSpace;
    public static $ZeroWidthNonJoiner;
    public static $ZeroWidthJoiner;
    public static $LeftToRightMark;
    public static $RightToLeftMark;
    public static $PopDirectionalFormatting;
    public static $LeftToRightOverride;
    public static $RightToLeftOverride;
    public static $NarrowNoBreakSpace;
    public static $DottedCircle;
    public static $OrnateLeftParenthesis;
    public static $OrnateRightParenthesis;
    public static function MakeUniRegEx($input)
    {
        return implode(array_map(explode($input), function($ch) { return "\\u" . substr("0000" . bin2hex($ch), -4); }));
    }
    public static function MakeRegMultiEx($input)
	{
        return implode("|", $input);
    }
};

class CachedData
{
	public static $_IslamData = null;
	public static function IslamData() { if (CachedData::$_IslamData === null) CachedData::$_IslamData = simplexml_load_file(dirname(__FILE__) . "/metadata/islaminfo.xml"); return CachedData::$_IslamData; }
    static $_AlDuri = null;
    static $_Warsh = null;
    static $_WarshScript = null;
    static $_UthmaniMinimalScript = null;
    static $_SimpleEnhancedScript = null;
    static $_SimpleScript = null;
    static $_SimpleCleanScript = null;
    static $_SimpleMinimalScript = null;
    static $_RomanizationRules = null;
    static $_ColoringSpelledOutRules = null;
    static $_ErrorCheck = null;
    static $_RulesOfRecitationRegEx = null;
    public static function GetNum($name)
    {
        for ($count = 0; $count < count(CachedData::IslamData()->ArabicNumbers); $count++) {
            if (CachedData::IslamData()->ArabicNumbers[$count]->Name == $name) {
                return CachedData::IslamData()->ArabicNumbers[$count]->Text;
            }
        }
        return [];
    }
    public static function GetPattern($name)
    {
        for ($count = 0; $count < count(CachedData::IslamData()->ArabicPatterns); $count++) {
            if (CachedData::IslamData()->ArabicPatterns[$count]->Name == $name) {
                return CachedData::TranslateRegEx(CachedData::IslamData()->ArabicPatterns[$count]->Match, true);
            }
        }
        return "";
    }
    public static function GetGroup($name)
    {
        for ($count = 0; $count < count(CachedData::IslamData()->ArabicGroups); $count++) {
            if (CachedData::IslamData()->ArabicGroups[$count]->Name == $name) {
                return array_map(CachedData::IslamData()->ArabicGroups[$count]->Text, function($str) { return CachedData::TranslateRegEx($str, $name == "ArabicSpecialLetters"); });
            }
        }
        return [];
    }
    public static function GetRuleSet($name)
    {
        for ($count = 0; $count < count(CachedData::IslamData()->RuleSets); $count++) {
            if (CachedData::IslamData()->RuleSets[$count]->Name == $name) {
                $ruleSet = CachedData::IslamData()->RuleSets[$count]->Rules;
                for ($subCount = 0; $subCount < count($ruleSet); $subCount++) {
                    $ruleSet[$subCount]->Match = CachedData::TranslateRegEx($ruleSet[$subCount]->Match, true);
                    $ruleSet[$subCount]->Evaluator = CachedData::TranslateRegEx($ruleSet[$subCount]->Evaluator, false);
                }
                return $ruleSet;
            }
        }
        return null;
    }
    static $_ArabicUniqueLetters = null;
    static $_ArabicNumbers = null;
    static $_ArabicWaslKasraExceptions = null;
    static $_ArabicBaseNumbers = null;
    static $_ArabicBaseExtraNumbers = null;
    static $_ArabicBaseTenNumbers = null;
    static $_ArabicBaseHundredNumbers = null;
    static $_ArabicBaseThousandNumbers = null;
    static $_ArabicBaseMillionNumbers = null;
    static $_ArabicBaseMilliardNumbers = null;
    static $_ArabicBaseBillionNumbers = null;
    static $_ArabicBaseTrillionNumbers = null;
    static $_ArabicFractionNumbers = null;
    static $_ArabicOrdinalNumbers = null;
    static $_ArabicOrdinalExtraNumbers = null;
    public static function ArabicUniqueLetters()
	{
        if (CachedData::$_ArabicUniqueLetters == null) {
            CachedData::$_ArabicUniqueLetters = CachedData::GetNum("ArabicUniqueLetters");
        }
    	return CachedData::$_ArabicUniqueLetters;
    }
    public static function ArabicNumbers()
	{
        if (CachedData::$_ArabicNumbers == null) {
            CachedData::$_ArabicNumbers = CachedData::GetNum("ArabicNumbers");
        }
        return CachedData::$_ArabicNumbers;
    }
    public static function ArabicWaslKasraExceptions()
	{
        if (CachedData::$_ArabicWaslKasraExceptions == null) {
            CachedData::$_ArabicWaslKasraExceptions = CachedData::GetNum("ArabicWaslKasraExceptions");
        }
        return CachedData::$_ArabicWaslKasraExceptions;
    }
    public static function ArabicBaseNumbers()
	{
        if (CachedData::$_ArabicBaseNumbers == null) {
            CachedData::$_ArabicBaseNumbers = CachedData::GetNum("base");
        }
        return CachedData::$_ArabicBaseNumbers;
    }
    public static function ArabicBaseExtraNumbers()
	{
        if (CachedData::$_ArabicBaseExtraNumbers == null) {
            CachedData::$_ArabicBaseExtraNumbers = CachedData::GetNum("baseextras");
        }
        return CachedData::$_ArabicBaseExtraNumbers;
    }
    public static function ArabicBaseTenNumbers()
	{
        if (CachedData::$_ArabicBaseTenNumbers == null) {
            CachedData::$_ArabicBaseTenNumbers = CachedData::GetNum("baseten");
        }
        return CachedData::$_ArabicBaseTenNumbers;
    }
    public static function ArabicBaseHundredNumbers()
	{
        if (CachedData::$_ArabicBaseHundredNumbers == null) {
            CachedData::$_ArabicBaseHundredNumbers = CachedData::GetNum("basehundred");
        }
        return CachedData::$_ArabicBaseHundredNumbers;
    }
    public static function ArabicBaseThousandNumbers()
	{
        if (CachedData::$_ArabicBaseThousandNumbers == null) {
            CachedData::$_ArabicBaseThousandNumbers = CachedData::GetNum("thousands");
        }
        return CachedData::$_ArabicBaseThousandNumbers;
    }
    public static function ArabicBaseMillionNumbers()
	{
        if (CachedData::$_ArabicBaseMillionNumbers == null) {
            CachedData::$_ArabicBaseMillionNumbers = CachedData::GetNum("millions");
        }
        return CachedData::$_ArabicBaseMillionNumbers;
    }
    public static function ArabicBaseMilliardNumbers()
	{
        if (CachedData::$_ArabicBaseMilliardNumbers == null) {
            CachedData::$_ArabicBaseMilliardNumbers = CachedData::GetNum("milliard");
        }
        return CachedData::$_ArabicBaseMilliardNumbers;
    }
    public static function ArabicBaseBillionNumbers()
	{
        if (CachedData::$_ArabicBaseBillionNumbers == null) {
            CachedData::$_ArabicBaseBillionNumbers = CachedData::GetNum("billions");
        }
        return CachedData::$_ArabicBaseBillionNumbers;
    }
    public static function ArabicBaseTrillionNumbers()
	{
        if (CachedData::$_ArabicBaseTrillionNumbers == null) {
            CachedData::$_ArabicBaseTrillionNumbers = CachedData::GetNum("trillions");
        }
        return CachedData::$_ArabicBaseTrillionNumbers;
    }
    public static function ArabicFractionNumbers()
	{
        if (CachedData::$_ArabicFractionNumbers == null) {
            CachedData::$_ArabicFractionNumbers = CachedData::GetNum("fractions");
        }
        return CachedData::$_ArabicFractionNumbers;
    }
    public static function ArabicOrdinalNumbers()
	{
        if (CachedData::$_ArabicOrdinalNumbers == null) {
            CachedData::$_ArabicOrdinalNumbers = CachedData::GetNum("ordinals");
        }
        return CachedData::$_ArabicOrdinalNumbers;
    }
    public static function ArabicOrdinalExtraNumbers()
	{
        if (CachedData::$_ArabicOrdinalExtraNumbers == null) {
            CachedData::$_ArabicOrdinalExtraNumbers = CachedData::GetNum("ordinalextras");
        }
        return CachedData::$_ArabicOrdinalExtraNumbers;
    }
    static $_CertainStopPattern = null;
    static $_OptionalStopPattern = null;
    static $_OptionalStopPatternNotEndOfAyah = null;
    static $_CertainNotStopPattern = null;
    static $_OptionalNotStopPattern = null;
    static $_TehMarbutaStopRule = null;
    static $_TehMarbutaContinueRule = null;
    public static function CertainStopPattern()
	{
        if (CachedData::$_CertainStopPattern == null) {
            CachedData::$_CertainStopPattern = CachedData::GetPattern("CertainStopPattern");
        }
        return CachedData::$_CertainStopPattern;
    }
    public static function OptionalStopPattern()
	{
        if (CachedData::$_OptionalStopPattern == null) {
            CachedData::$_OptionalStopPattern = CachedData::GetPattern("OptionalStopPattern");
        }
        return CachedData::$_OptionalStopPattern;
    }
    public static function OptionalStopPatternNotEndOfAyah()
	{
        if (CachedData::$_OptionalStopPatternNotEndOfAyah == null) {
            CachedData::$_OptionalStopPatternNotEndOfAyah = CachedData::GetPattern("OptionalStopPatternNotEndOfAyah");
        }
        return CachedData::$_OptionalStopPatternNotEndOfAyah;
    }
    public static function CertainNotStopPattern()
	{
        if (CachedData::$_CertainNotStopPattern == null) {
            CachedData::$_CertainNotStopPattern = CachedData::GetPattern("CertainNotStopPattern");
        }
        return CachedData::$_CertainNotStopPattern;
    }
    public static function OptionalNotStopPattern()
	{
        if (CachedData::$_OptionalNotStopPattern == null) {
            CachedData::$_OptionalNotStopPattern = CachedData::GetPattern("OptionalNotStopPattern");
        }
        return CachedData::$_OptionalNotStopPattern;
    }
    public static function TehMarbutaStopRule()
	{
        if (CachedData::$_TehMarbutaStopRule == null) {
            CachedData::$_TehMarbutaStopRule = CachedData::GetPattern("TehMarbutaStopRule");
        }
        return CachedData::$_TehMarbutaStopRule;
    }
    public static function TehMarbutaContinueRule()
	{
        if (CachedData::$_TehMarbutaContinueRule == null) {
            CachedData::$_TehMarbutaContinueRule = CachedData::GetPattern("TehMarbutaContinueRule");
        }
        return CachedData::$_TehMarbutaContinueRule;
    }
    static $_ArabicLongVowels = null;
    static $_ArabicTanweens = null;
    static $_ArabicFathaDammaKasra = null;
    static $_ArabicStopLetters = null;
    static $_ArabicSpecialGutteral = null;
    static $_ArabicSpecialLeadingGutteral = null;
    static $_ArabicPunctuationSymbols = null;
    static $_ArabicLetters = null;
    static $_ArabicSunLettersNoLam = null;
    static $_ArabicSunLetters = null;
    static $_ArabicMoonLettersNoVowels = null;
    static $_ArabicMoonLetters = null;
    static $_RecitationCombiningSymbols = null;
    static $_RecitationConnectingFollowerSymbols = null;
    static $_RecitationSymbols = null;
    static $_ArabicLettersInOrder = null;
    static $_ArabicSpecialLetters = null;
    static $_ArabicHamzas = null;
    static $_ArabicVowels = null;
    static $_ArabicTajweed = null;
    static $_ArabicPunctuation = null;
    static $_NonArabicLetters = null;
    static $_WhitespaceSymbols = null;
    static $_PunctuationSymbols = null;
    static $_RecitationDiacritics = null;
    static $_RecitationLettersDiacritics = null;
    static $_RecitationSpecialSymbols = null;
    static $_ArabicLeadingGutterals = null;
    static $_RecitationLetters = null;
    static $_ArabicTrailingGutterals = null;
    static $_RecitationSpecialSymbolsNotStop = null;
    public static function ArabicLongVowels()
	{
        if (CachedData::$_ArabicLongVowels == null) {
            CachedData::$_ArabicLongVowels = CachedData::GetGroup("ArabicLongVowels");
        }
        return CachedData::$_ArabicLongVowels;
    }
    public static function ArabicTanweens()
	{
        if (CachedData::$_ArabicTanweens == null) {
            CachedData::$_ArabicTanweens = CachedData::GetGroup("ArabicTanweens");
        }
        return CachedData::$_ArabicTanweens;
    }
    public static function ArabicFathaDammaKasra()
	{
        if (CachedData::$_ArabicFathaDammaKasra == null) {
            CachedData::$_ArabicFathaDammaKasra = CachedData::GetGroup("ArabicFathaDammaKasra");
        }
        return CachedData::$_ArabicFathaDammaKasra;
    }
    public static function ArabicStopLetters()
	{
        if (CachedData::$_ArabicStopLetters == null) {
            CachedData::$_ArabicStopLetters = CachedData::GetGroup("ArabicStopLetters");
        }
        return CachedData::$_ArabicStopLetters;
    }
    public static function ArabicSpecialGutteral()
	{
        if (CachedData::$_ArabicSpecialGutteral == null) {
            CachedData::$_ArabicSpecialGutteral = CachedData::GetGroup("ArabicSpecialGutteral");
        }
        return CachedData::$_ArabicSpecialGutteral;
    }
    public static function ArabicSpecialLeadingGutteral()
	{
        if (CachedData::$_ArabicSpecialLeadingGutteral == null) {
            CachedData::$_ArabicSpecialLeadingGutteral = CachedData::GetGroup("ArabicSpecialLeadingGutteral");
        }
        return CachedData::$_ArabicSpecialLeadingGutteral;
    }
    public static function ArabicPunctuationSymbols()
	{
        if (CachedData::$_ArabicPunctuationSymbols == null) {
            CachedData::$_ArabicPunctuationSymbols = CachedData::GetGroup("ArabicPunctuationSymbols");
        }
        return CachedData::$_ArabicPunctuationSymbols;
    }
    public static function ArabicLetters()
	{
        if (CachedData::$_ArabicLetters == null) {
            CachedData::$_ArabicLetters = CachedData::GetGroup("ArabicLetters");
        }
        return CachedData::$_ArabicLetters;
    }
    public static function ArabicSunLettersNoLam()
	{
        if (CachedData::$_ArabicSunLettersNoLam == null) {
            CachedData::$_ArabicSunLettersNoLam = CachedData::GetGroup("ArabicSunLettersNoLam");
        }
        return CachedData::$_ArabicSunLettersNoLam;
    }
    public static function ArabicSunLetters()
	{
        if (CachedData::$_ArabicSunLetters == null) {
            CachedData::$_ArabicSunLetters = CachedData::GetGroup("ArabicSunLetters");
        }
        return CachedData::$_ArabicSunLetters;
    }
    public static function ArabicMoonLettersNoVowels()
	{
        if (CachedData::$_ArabicMoonLettersNoVowels == null) {
            CachedData::$_ArabicMoonLettersNoVowels = CachedData::GetGroup("ArabicMoonLettersNoVowels");
        }
        return CachedData::$_ArabicMoonLettersNoVowels;
    }
    public static function ArabicMoonLetters()
	{
        if (CachedData::$_ArabicMoonLetters == null) {
            CachedData::$_ArabicMoonLetters = CachedData::GetGroup("ArabicMoonLetters");
        }
        return CachedData::$_ArabicMoonLetters;
    }
    public static function RecitationCombiningSymbols()
	{
        if (CachedData::$_RecitationCombiningSymbols == null) {
            CachedData::$_RecitationCombiningSymbols = CachedData::GetGroup("RecitationCombiningSymbols");
        }
        return CachedData::$_RecitationCombiningSymbols;
    }
    public static function RecitationConnectingFollowerSymbols()
	{
        if (CachedData::$_RecitationConnectingFollowerSymbols == null) {
            CachedData::$_RecitationConnectingFollowerSymbols = CachedData::GetGroup("RecitationConnectingFollowerSymbols");
        }
        return CachedData::$_RecitationConnectingFollowerSymbols;
    }
    public static function RecitationSymbols()
	{
        if (CachedData::$_RecitationSymbols == null) {
            CachedData::$_RecitationSymbols = CachedData::GetGroup("RecitationSymbols");
        }
        return CachedData::$_RecitationSymbols;
    }
    public static function ArabicLettersInOrder()
	{
        if (CachedData::$_ArabicLettersInOrder == null) {
            CachedData::$_ArabicLettersInOrder = CachedData::GetGroup("ArabicLettersInOrder");
        }
        return CachedData::$_ArabicLettersInOrder;
    }
    public static function ArabicSpecialLetters()
	{
        if (CachedData::$_ArabicSpecialLetters == null) {
            CachedData::$_ArabicSpecialLetters = CachedData::GetGroup("ArabicSpecialLetters");
        }
        return CachedData::$_ArabicSpecialLetters;
    }
    public static function ArabicHamzas()
	{
        if (CachedData::$_ArabicHamzas == null) {
            CachedData::$_ArabicHamzas = CachedData::GetGroup("ArabicHamzas");
        }
        return CachedData::$_ArabicHamzas;
    }
    public static function ArabicVowels()
	{
        if (CachedData::$_ArabicVowels == null) {
            CachedData::$_ArabicVowels = CachedData::GetGroup("ArabicVowels");
        }
        return CachedData::$_ArabicVowels;
    }
    public static function ArabicTajweed()
	{
        if (CachedData::$_ArabicTajweed == null) {
            CachedData::$_ArabicTajweed = CachedData::GetGroup("ArabicTajweed");
        }
        return CachedData::$_ArabicTajweed;
    }
    public static function ArabicPunctuation()
	{
        if (CachedData::$_ArabicPunctuation == null) {
            CachedData::$_ArabicPunctuation = CachedData::GetGroup("ArabicPunctuation");
        }
        return CachedData::$_ArabicPunctuation;
    }
    public static function NonArabicLetters()
	{
        if (CachedData::$_NonArabicLetters == null) {
            CachedData::$_NonArabicLetters = CachedData::GetGroup("NonArabicLetters");
        }
        return CachedData::$_NonArabicLetters;
    }
    public static function WhitespaceSymbols()
	{
        if (CachedData::$_WhitespaceSymbols == null) {
            CachedData::$_WhitespaceSymbols = CachedData::GetGroup("WhitespaceSymbols");
        }
        return CachedData::$_WhitespaceSymbols;
    }
    public static function PunctuationSymbols()
	{
        if (CachedData::$_PunctuationSymbols == null) {
            CachedData::$_PunctuationSymbols = CachedData::GetGroup("PunctuationSymbols");
        }
        return CachedData::$_PunctuationSymbols;
    }
    public static function RecitationDiacritics()
	{
        if (CachedData::$_RecitationDiacritics == null) {
            CachedData::$_RecitationDiacritics = CachedData::GetGroup("RecitationDiacritics");
        }
        return CachedData::$_RecitationDiacritics;
    }
    public static function RecitationLettersDiacritics()
	{
        if (CachedData::$_RecitationLettersDiacritics == null) {
            CachedData::$_RecitationLettersDiacritics = CachedData::GetGroup("RecitationLettersDiacritics");
        }
        return CachedData::$_RecitationLettersDiacritics;
    }
    public static function RecitationSpecialSymbols()
	{
        if (CachedData::$_RecitationSpecialSymbols == null) {
            CachedData::$_RecitationSpecialSymbols = CachedData::GetGroup("RecitationSpecialSymbols");
        }
        return CachedData::$_RecitationSpecialSymbols;
    }
    public static function ArabicLeadingGutterals()
	{
        if (CachedData::$_ArabicLeadingGutterals == null) {
            CachedData::$_ArabicLeadingGutterals = CachedData::GetGroup("ArabicLeadingGutterals");
        }
        return CachedData::$_ArabicLeadingGutterals;
    }
    public static function RecitationLetters()
	{
        if (CachedData::$_RecitationLetters == null) {
            CachedData::$_RecitationLetters = CachedData::GetGroup("RecitationLetters");
        }
        return CachedData::$_RecitationLetters;
    }
    public static function ArabicTrailingGutterals()
	{
        if (CachedData::$_ArabicTrailingGutterals == null) {
            CachedData::$_ArabicTrailingGutterals = CachedData::GetGroup("ArabicTrailingGutterals");
        }
        return CachedData::$_ArabicTrailingGutterals;
    }
    public static function RecitationSpecialSymbolsNotStop()
	{
        if (CachedData::$_RecitationSpecialSymbolsNotStop == null) {
            CachedData::$_RecitationSpecialSymbolsNotStop = CachedData::GetGroup("RecitationSpecialSymbolsNotStop");
        }
        return CachedData::$_RecitationSpecialSymbolsNotStop;
    }
    public static function TranslateRegEx($value, $bAll)
    {
        return preg_replace_callback($value, "\{(.*?)\}", function($matches) {
            if (bAll) {
                if (Match.Groups(1).Value = "CertainStopPattern") { return CachedData::CertainStopPattern(); }
                if (Match.Groups(1).Value = "OptionalStopPattern") { return CachedData::OptionalStopPattern(); }
                if (Match.Groups(1).Value = "OptionalStopPatternNotEndOfAyah") { return CachedData::OptionalStopPatternNotEndOfAyah(); }
                if (Match.Groups(1).Value = "CertainNotStopPattern") { return CachedData::CertainNotStopPattern(); }
                if (Match.Groups(1).Value = "OptionalNotStopPattern") { return CachedData::OptionalNotStopPattern(); }
                if (Match.Groups(1).Value = "TehMarbutaStopRule") { return CachedData::TehMarbutaStopRule(); }
                if (Match.Groups(1).Value = "TehMarbutaContinueRule") { return CachedData::TehMarbutaContinueRule(); }

                if (Match.Groups(1).Value = "ArabicUniqueLetters") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicUniqueLetters, Function($str) ArabicData.TransliterateFromBuckwalter(Str).Replace(CStr(ArabicData.ArabicMaddahAbove), String.Empty))); }
                if (Match.Groups(1).Value = "ArabicNumbers") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicNumbers, Function($str) ArabicData.TransliterateFromBuckwalter(Str))); }
                if (Match.Groups(1).Value = "ArabicWaslKasraExceptions") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicWaslKasraExceptions, Function($str) ArabicData.TransliterateFromBuckwalter(Str))); }
                //if (Match.Groups(1).Value = "SimpleSuperscriptAlefBefore") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(Arabic.SimpleSuperscriptAlefBefore, Function($str) ArabicData.TransliterateFromBuckwalter(Str.Replace(".", String.Empty).Replace("""", String.Empty).Replace("@", String.Empty).Replace("[", String.Empty).Replace("]", String.Empty).Replace("-", String.Empty).Replace("^", String.Empty)))); }
                //if (Match.Groups(1).Value = "SimpleSuperscriptAlefNotBefore") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(Arabic.SimpleSuperscriptAlefNotBefore, Function($str) ArabicData.TransliterateFromBuckwalter(Str.Replace(".", String.Empty).Replace("""", String.Empty).Replace("@", String.Empty).Replace("[", String.Empty).Replace("]", String.Empty).Replace("-", String.Empty).Replace("^", String.Empty)))); }
                //if (Match.Groups(1).Value = "SimpleSuperscriptAlefAfter") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(Arabic.SimpleSuperscriptAlefAfter, Function($str) ArabicData.TransliterateFromBuckwalter(Str.Replace(".", String.Empty).Replace("""", String.Empty).Replace("@", String.Empty).Replace("[", String.Empty).Replace("]", String.Empty).Replace("-", String.Empty).Replace("^", String.Empty)))); }
                //if (Match.Groups(1).Value = "SimpleSuperscriptAlefNotAfter") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(Arabic.SimpleSuperscriptAlefNotAfter, Function($str) ArabicData.TransliterateFromBuckwalter(Str.Replace(".", String.Empty).Replace("""", String.Empty).Replace("@", String.Empty).Replace("[", String.Empty).Replace("]", String.Empty).Replace("-", String.Empty).Replace("^", String.Empty)))); }
                if (Match.Groups(1).Value = "ArabicLongShortVowels") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicLongVowels, Function($strV) ArabicData.MakeUniRegEx(StrV(0) + "(?=" + ArabicData.MakeUniRegEx(StrV(1)) + ")"))); }
                if (Match.Groups(1).Value = "ArabicTanweens") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "ArabicFathaDammaKasra") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "ArabicStopLetters") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicStopLetters, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "ArabicSpecialGutteral") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicSpecialGutteral, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "ArabicSpecialLeadingGutteral") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "ArabicPunctuationSymbols") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicPunctuationSymbols, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "ArabicLetters") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "ArabicSunLettersNoLam") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicSunLettersNoLam, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "ArabicSunLetters") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "ArabicMoonLettersNoVowels") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicMoonLettersNoVowels, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "ArabicMoonLetters") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicMoonLetters, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "RecitationCombiningSymbols") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(RecitationCombiningSymbols, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "RecitationConnectingFollowerSymbols") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(RecitationConnectingFollowerSymbols, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "PunctuationSymbols") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(PunctuationSymbols, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "RecitationLettersDiacritics") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(RecitationLettersDiacritics, Function($str) ArabicData.MakeUniRegEx(Str))); }
                if (Match.Groups(1).Value = "RecitationSpecialSymbolsNotStop") { return ArabicData.MakeRegMultiEx(Array.ConvertAll(RecitationSpecialSymbolsNotStop, Function($str) ArabicData.MakeUniRegEx(Str))); }
            }
            if (preg_match(Match.Groups(1).Value, "0x([0-9a-fA-F]{4})").Success) {
                return bAll ? ArabicData.MakeUniRegEx(ChrW(Integer.Parse(Match.Groups(1).Value.Substring(2), System.Globalization.NumberStyles.HexNumber))) : ChrW(Integer.Parse(Match.Groups(1).Value.Substring(2), System.Globalization.NumberStyles.HexNumber)));
            }
            for ($count = 0; $count < count(ArabicData::Data::ArabicLetters); $count++) {
                if (Match.Groups(1).Value = ArabicData.Data.ArabicLetters(Count).UnicodeName) { return If(bAll, ArabicData.MakeUniRegEx(ArabicData.Data.ArabicLetters(Count).Symbol), ArabicData.Data.ArabicLetters(Count).Symbol)
            }
           	//{0} ignore
            return Match.Value;
        });
    }
    public static function RulesOfRecitationRegEx()
	{
        if (CachedData::$_RulesOfRecitationRegEx == null) {
            CachedData::$_RulesOfRecitationRegEx = CachedData::IslamData()->MetaRules;
            for ($subCount = 0; $subCount < count(_RulesOfRecitationRegEx); $subCount++) {
                CachedData::$_RulesOfRecitationRegEx[$subCount]->Match = TranslateRegEx(CachedData::$_RulesOfRecitationRegEx[$subCount]->Match, true);
            }
        }
        return CachedData::$_RulesOfRecitationRegEx;
    }
    public static function WarshScript()
	{
        if (CachedData::$_WarshScript == null) {
            CachedData::$_WarshScript = CachedData::GetRuleSet("WarshScript");
        }
        return CachedData::$_WarshScript;
    }
    public static function UthmaniMinimalScript()
	{
        if (CachedData::$_UthmaniMinimalScript == null) {
            CachedData::$_UthmaniMinimalScript = CachedData::GetRuleSet("UthmaniMinimalScript");
        }
        return CachedData::$_UthmaniMinimalScript;
    }
    public static function SimpleEnhancedScript()
	{
        if (CachedData::$_SimpleEnhancedScript == null) {
            CachedData::$_SimpleEnhancedScript = CachedData::GetRuleSet("SimpleEnhancedScript");
        }
        return CachedData::$_SimpleEnhancedScript;
    }
    public static function SimpleScript()
	{
        if (CachedData::$_SimpleScript == null) {
            CachedData::$_SimpleScript = CachedData::GetRuleSet("SimpleScript");
        }
        return CachedData::$_SimpleScript;
    }
    public static function SimpleCleanScript()
	{
        if (CachedData::$_SimpleCleanScript == null) {
            CachedData::$_SimpleCleanScript = CachedData::GetRuleSet("SimpleCleanScript");
        }
        return CachedData::$_SimpleCleanScript;
    }
    public static function SimpleMinimalScript()
	{
        if (CachedData::$_SimpleMinimalScript == null) {
            CachedData::$_SimpleMinimalScript = CachedData::GetRuleSet("SimpleMinimalScript");
        }
        return CachedData::$_SimpleMinimalScript;
    }
    public static function RomanizationRules()
	{
        if (CachedData::$_RomanizationRules == null) {
            CachedData::$_RomanizationRules = CachedData::GetRuleSet("RomanizationRules");
        }
        return CachedData::$_RomanizationRules;
    }
    public static function ColoringSpelledOutRules()
	{
        if (CachedData::$_ColoringSpelledOutRules == null) {
            CachedData::$_ColoringSpelledOutRules = CachedData::GetRuleSet("ColoringSpelledOutRules");
        }
        return CachedData::$_ColoringSpelledOutRules;
    }
    public static function ErrorCheckRules()
	{
        if (CachedData::$_ErrorCheck == null) {
            CachedData::$_ErrorCheck = CachedData::GetRuleSet("ErrorCheck");
        }
        return CachedData::$_ErrorCheck;
    }
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
	};
	abstract class RenderDisplayClass {
		const eNested = 0;
		const eArabic = 1;
		const eTransliteration = 2;
		const eLTR = 3;
		const eRTL = 4;
		const eContinueStop = 5;
		const eRanking = 6;
		const eList = 7;
	};
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
        for ($count = 0; $count < count(CachedData::IslamData()->TranslitSchemes); $count++) {
            if (CachedData::IslamData()->TranslitSchemes[$count]->Name == $scheme) {
                $sch = CachedData::IslamData()->TranslitSchemes[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData()->TranslitSchemes) { return ""; }
        return $sch->SpecialLetters[$index];
    }
    public static function GetSchemeSpecialFromMatch($str, $scheme, $bExp)
    {
    	$sch = null;
        for ($count = 0; $count < count(CachedData::IslamData()->TranslitSchemes); $count++) {
            if (CachedData::IslamData()->TranslitSchemes[$count]->Name == $scheme) {
                $sch = CachedData::IslamData()->TranslitSchemes[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData()->TranslitSchemes)) { return -1; }
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
        for ($count = 0; $count < count(CachedData::IslamData()->TranslitSchemes); $count++) {
            if (CachedData::IslamData()->TranslitSchemes[$count]->Name == $scheme) {
                $sch = CachedData::IslamData()->TranslitSchemes[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData()->TranslitSchemes)) { return ""; }
        if (array_search($str, CachedData::ArabicVowels) !== false) {
            return $sch->Vowels(array_search($str, CachedData::ArabicVowels));
        }
        return "";
    }
    public static function GetSchemeGutteralFromString($str, $scheme, $leading)
    {
        $sch = null;
        for ($count = 0; $count < count(CachedData::IslamData()->TranslitSchemes); $count++) {
            if (CachedData::IslamData()->TranslitSchemes[$count]->Name == $scheme) {
                $sch = CachedData::IslamData()->TranslitSchemes[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData()->TranslitSchemes)) { return ""; }
        if (array_search($str, CachedData::ArabicLeadingGutterals) !== false) {
            return $sch->Vowels(array_search($str, CachedData::ArabicLeadingGutterals) + count(CachedData::ArabicVowels) + ($leading ? count(CachedData::ArabicLeadingGutterals) : 0));
        }
        return "";
    }
    public static function GetSchemeValueFromSymbol($symbol, $scheme)
    {
        $sch = null;
        for ($count = 0; $count < count(CachedData::IslamData()->TranslitSchemes); $count++) {
            if (CachedData::IslamData()->TranslitSchemes[$count]->Name == $scheme) {
                $sch = CachedData::IslamData()->TranslitSchemes[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData()->TranslitSchemes)) { return ""; }
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
	abstract class RuleFuncs
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
        $output = "";
        foreach ($ch => $input) {
            $index = ArabicData::FindLetterBySymbol($ch);
            if ($index != -1 && Arabic::IsLetter($index)) {
                if ($output != "" && !$quranic) { $output .= " "; }
                Output .= $quranic ? ArabicData::Data::ArabicLetters[$index]->SymbolName.Remove(strlen(ArabicData::Data::ArabicLetters[$index]->SymbolName) - 1) + (ArabicData::Data::ArabicLetters[$index]->SymbolName.EndsWith("n") ? "" : "o"), ArabicData::Data::ArabicLetters[$index]->SymbolName);
            } elseif ($index != -1 && ArabicData::Data::ArabicLetters[$index]->Symbol == ArabicData::ArabicMaddahAbove) {
                if (!$quranic) { $output += $ch; }
            }
        }
        return ArabicData::TransliterateFromBuckwalter($output);
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
        $str = "";
        $hundStr = "";
        if ($number >= 100) {
            $hundStr = $useAlefHundred ? CachedData::ArabicBaseHundredNumbers(($number \ 100) - 1).Insert(2, "A") : CachedData::ArabicBaseHundredNumbers(($number \ 100) - 1);
            if ((Number % 100) = 0) { return $hundStr; }
            $number = $number % 100;
        }
        if (($number % 10) != 0 && $number != 11 && $number != 12) {
            $str = CachedData::ArabicBaseNumbers(($number % 10) - 1);
        }
        if ($number >= 11 && $number < 20) {
            if ($number == 11 || $number == 12) {
                $str .= CachedData::ArabicBaseExtraNumbers($number - 11);
            Else
                $str = Str.Remove(Str.Length - 1) . "a";
            }
            $str .= " Ea$ara";
        } elseif (($number == 0 && $str == "") || $number == 10 || &number >= 20) {
            $str = ($str == "" ? "" : $str . " wa") . CachedData::ArabicBaseTenNumbers($number \ 10);
        }
        return $useClassic ? ($str == "" ? "" : $str . ($hundStr == "", "", " wa")) . $hundStr : ($hundStr == "" ? "" : $hundStr . ($str == "" ? "" : " wa")) . $str;
    }
    public static function ArabicWordFromNumber($number, $useClassic, $useAlefHundred, $useMilliard)
    {
        $str = "";
        $nextStr = "";
        $curBase = 3;
        $baseNums = [1000, 1000000, 1000000000, 1000000000000];
        $bases = [CachedData::ArabicBaseThousandNumbers, CachedData::ArabicBaseMillionNumbers, $useMilliard ? CachedData::ArabicBaseMilliardNumbers : CachedData::ArabicBaseBillionNumbers, CachedData::ArabicBaseTrillionNumbers];
        do {
            if ($number >= $baseNums[$curBase] && $number < 2 * $baseNums[$curBase]) {
                $nextStr = $bases[$curBase][0];
            } elseif ($number >= 2 * $baseNums[$curBase] && $number < 3 * $baseNums[$curBase]) {
                $nextStr = $bases[$curBase][1];
            } elseif ($number >= 3 * $baseNums[$curBase] && $number < 10 * $baseNums[$curBase]) {
                $nextStr = CachedData::ArabicBaseNumbers(CInt($number \ $baseNums[$curBase] - 1)).Remove(CachedData::ArabicBaseNumbers(CInt($number \ $baseNums[$curBase] - 1)).Length - 1) . "u " . $bases[$curBase][2];
            } elseif ($number >= 10 * $baseNums[$curBase] && $number < 11 * $baseNums[$curBase]) {
                $nextStr = CachedData::ArabicBaseTenNumbers[1].Remove(strlen(CachedData::ArabicBaseTenNumbers[1]) - 1) . "u " . $bases[$curBase][2];
            } elseif ($number >= $baseNums[$curBase]) {
                $nextStr = Arabic::ArabicWordForLessThanThousand(CInt(($number \ $baseNums[$curBase]) % 100), $useClassic, $useAlefHundred);
                if ($number >= 100 * $baseNums[$curBase] && $number < If(UseClassic, 200, 101) * $baseNums[$curBase]) {
                    $nextStr = $nextStr.Remove(strlen($nextStr) - 1) + "u " + $bases[$curBase][0].Remove(strlen($bases[$curBase][0]) - 1) + "K"
                } elseif ($number >= 200 * $baseNums[$curBase] && $number < If($useClassic, 300, 201) * $baseNums[$curBase]) {
                    $nextStr = $nextStr.Remove(strlen($nextStr) - 2) + " " + $bases[$curBase][0].Remove(strlen($bases[$curBase][0]) - 1) + "K"
                } elseif ($number >= 300 * $baseNums[$curBase] && ($useClassic Or ($number \ $baseNums[$curBase]) Mod 100 = 0)) {
                    $nextStr = $nextStr.Remove(strlen($nextStr) - 1) + "i " + $bases[$curBase][0].Remove(strlen($bases[$curBase][0]) - 1) + "K"
                } else {
                    $nextStr .= " " . $bases[$curBase][0].Remove(strlen($bases[$curBase][0]) - 1) + "FA"
                }
            }
            $number = $number % $baseNums[$curBase];
            $curBase -= 1;
            $str = $useClassic ? ($nextStr == "" ? "" : $nextStr . ($str == "" ? "" : " wa")) . $str : ($str == "" ? "" : $str + ($nextStr == "" ? "" : " wa")) . $nextStr;
            $nextStr = "";
        } while ($curBase >= 0);
        if ($number <> 0 || $str == "") { $nextStr = Arabic::ArabicWordForLessThanThousand(CInt($number), $useClassic, $useAlefHundred); }
        return $useClassic ? ($nextStr == "" ? "" : $nextStr + ($str == "" ? "" : " wa")) . $str, ($str == "" ? "" : $str . ($nextStr == "" ? "" : " wa")) . $nextStr;
    }
    public static function ReplaceMetadata($arabicString, $metadataRule, $scheme, $optionalStops)
    {
        for ($count = 0; $count < count(CachedData.ColoringSpelledOutRules); $count++) {
            $match = Array.Find(CachedData.ColoringSpelledOutRules(Count).Match.Split("|"c), Function($str) Array.IndexOf(Array.ConvertAll(MetadataRule.Type.Split("|"c), Function($s) System.Text.RegularExpressions.Regex.Replace(S, "\(.*\)", String.Empty)), Str) <> -1)
            if ($match != null) {
                $str = String.Format(CachedData.ColoringSpelledOutRules(Count).Evaluator, ArabicString.Substring(MetadataRule.Index, MetadataRule.Length))
                if (CachedData.ColoringSpelledOutRules(Count).RuleFunc <> RuleFuncs.eNone) {
                    $args = RuleFunctions(CachedData.ColoringSpelledOutRules(Count).RuleFunc - 1)(Str, Scheme)
                    if (Args.Length = 1) {
                        $str = $args(0)
                    Else
                        $MetaArgs = System.Text.RegularExpressions.Regex.Match(MetadataRule.Type, Match + "\((.*)\)").Groups(1).Value.Split(","c)
                        $str = "";
                        for ($index = 0; $index < count($args); $index++) {
                            if ($args(Index) != null) {
                                Str += ReplaceMetadata(Args(Index), New RuleMetadata(0, Args(Index).Length, MetaArgs(Index).Replace(" "c, "|"c)), Scheme, OptionalStops)
                            }
                        }
                    }
                }
                $arabicString = $arabicString.Insert(MetadataRule.Index + MetadataRule.Length, Str).Remove(MetadataRule.Index, MetadataRule.Length)
            }
        }
        return $arabicString;
    }
    public static function DoErrorCheck($arabicString)
    {
        //need to check for decomposed first
        for ($count = 0; $count < count(CachedData::ErrorCheckRules); $count++) {
            $matches = preg_match_all($arabicString, CachedData::ErrorCheckRules[$count]->Match);
            for ($matchIndex = 0; $matchIndex < count($matches); $matchIndex++) {
                if (CachedData::ErrorCheckRules[$count]->NegativeMatch == null || $matches[$matchIndex]->Result(CachedData::ErrorCheckRules[$count]->NegativeMatch) == "") {
                    //Debug.Print(ErrorCheckRules[$count]->Rule . ": " . Arabic::TransliterateToScheme($arabicString, ArabicData::TranslitScheme::Literal, "").Insert($matches[$matchIndex]->Index, "<!-- -->"))
                }
            }
        }
    }
    public static function JoinContig($arabicString, $preString, $postString)
    {
        $index = $preString.LastIndexOf(" "c)
        //take last word of pre string and first word of post string or another if it is a pause marker
        //end of ayah sign without number is used as a proper place holder
        if ($index != -1 && strlen($preString) - 2 == $index) { $index = $preString.LastIndexOf(" "c, $index - 1); }
        if ($index != -1) { $preString = substr($preString, $index + 1); }
        if ($preString != "") { $preString .= " " + ArabicData::ArabicEndOfAyah . " "; }
        $index = $postString.IndexOf(" "c);
        if ($index == 2) { $index = $preString.IndexOf(" "c, $index + 1); }
        if ($index != -1) { $postString = substr($postString, 0, $index); }
        if ($postString != "") { $postString = " " . ArabicData::ArabicEndOfAyah . " " . $postString; }
        return $preString . $arabicString . $postString;
    }
    public static function UnjoinContig($arabicString, $preString, $postString)
    {
        $index = $arabicString.IndexOf(ArabicData::ArabicEndOfAyah)
        if (PreString <> String.Empty AndAlso Index <> -1) {
            ArabicString = ArabicString.Substring(Index + 1 + 1)
        }
        Index = ArabicString.LastIndexOf(ArabicData::ArabicEndOfAyah)
        if (PostString <> String.Empty AndAlso Index <> -1) {
            $arabicString = $arabicString.Substring(0, Index - 1);
        }
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
        for ($count = 0; $count < count(CachedData::IslamData()->TranslitSchemes); $count++) {
            $strings[$count * 2 + 2] = [Utility::LoadResourceString("IslamSource_" . CachedData::IslamData()->TranslitSchemes[$count]->Name), CStr($count * 2 + 2)];
            $strings[$count * 2 + 1 + 2] = [Utility::LoadResourceString("IslamSource_" . CachedData::IslamData()->TranslitSchemes[$count]->Name) . " Literal", CStr($count * 2 + 1 + 2)];
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