<?php
/*
Plugin Name: IslamSource
Description: Islam Source Quranic Verse and Islamic Phrase Plugin - Allows for Quranic chapters, verses even specified down through the word to be inserted easily by using formats {a:b:c-x:y:z} where b, c, x, y and z are optional depending on if a chapter, verse or particular word of a verse is desired or a range is desired so it could be in forms {a:b-y} or {a:b-x:y} such as the opening chapter which could be specified as {1:1-7}.  The Arabic is automatically displayed when posts are viewed.  It also allows for various calligraphy and Unicode Islamic words or phrases to be easily inserted through a button on the visual editor which are displayed when the posts are later viewed.
Version: 1.1.0
Author: IslamSource
Author URI: http://www.islamsource.info
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
    public static $FontList = ["AGAIslamicPhrases", "AGAArabesque", "Shia", "IslamicLogo", "KFGQPCArabicSymbols01", "Quranic", "Tulth", "Farsi", "Asmaul-Husna", "Asmaul-Husna_2"];
    public static $FontFile = ["AGA_Islamic_Phrases.TTF", "aga-arabesque.ttf", "SHIA.TTF", "islamic.ttf", "Symbols1_Ver02.otf", "Quranic.ttf", "Tulth.ttf", "Farsi.ttf", "Asmaul-Husna_1.ttf", "Asmaul-Husna_2.ttf"];
};

abstract class TranslitScheme
{
    const None = 0;
    const Literal = 1;
    const RuleBased = 2;
};

class ArabicData
{
    static $_ArabicXMLData = null;
    public static function Data()
    {
        if (ArabicData::$_ArabicXMLData == null) {
            ArabicData::$_ArabicXMLData = simplexml_load_file(dirname(__FILE__) . "/HostPage/metadata/ArabicData.xml");
        }
        return ArabicData::$_ArabicXMLData;
    }
	public static $_BuckwalterMap = null;
    public static function BuckwalterMap()
	{
        if (ArabicData::$_BuckwalterMap == null) {
            ArabicData::$_BuckwalterMap = array();
            for ($index = 0; $index < count(ArabicData::Data()->arabicletters->children()); $index++) {
                if ((string)ArabicData::Data()->arabicletters->children()[$index]->attributes()["extendedbuckwalter"] != "") {
                    ArabicData::$_BuckwalterMap[mb_convert_encoding('&#' . ArabicData::Data()->arabicletters->children()[$index]->attributes()["extendedbuckwalter"] . ';', 'UTF-8', 'HTML-ENTITIES')] = $index;
                }
            }
        }
        return ArabicData::$_BuckwalterMap;
    }
    public static function TransliterateFromBuckwalter($buckwalter)
    {
    	$arabicString = "";
        for ($count = 0; $count < strlen($buckwalter); $count++) {
            if ($buckwalter[$count] == "\\") {
                $count += 1;
                if ($buckwalter[$count] == ",") {
                    $arabicString .= ArabicData::$ArabicComma;
                } else {
                    $arabicString .= $buckwalter[$count];
                }
            } else {
                if (array_key_exists($buckwalter[$count], ArabicData::BuckwalterMap())) {
                    $arabicString .= mb_convert_encoding('&#x' . ArabicData::Data()->arabicletters->children()[ArabicData::BuckwalterMap()[$buckwalter[$count]]]->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES');
                } else {
                    $arabicString .= $buckwalter[$count];
                }
            }
        }
        return $arabicString;
    }
    public static $_ArabicLetterMap = null;
    public static function ArabicLetterMap()
    {
        if (ArabicData::$_ArabicLetterMap == null) {
            ArabicData::$_ArabicLetterMap = array();
            for ($index = 0; $index < count(ArabicData::Data()->arabicletters->children()); $index++) {
                if ((string)ArabicData::Data()->arabicletters->children()[$index]->attributes()["symbol"] != "") {
                    ArabicData::$_ArabicLetterMap[mb_convert_encoding('&#x' . ArabicData::Data()->arabicletters->children()[$index]->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES')] = $index;
                }
            }
        }
        return ArabicData::$_ArabicLetterMap;
	}
    public static function FindLetterBySymbol($symbol)
    {
        return array_key_exists($symbol, ArabicData::ArabicLetterMap()) ? ArabicData::ArabicLetterMap()[$symbol] : -1;
    }
    public static function init() {
        ArabicData::$Space = mb_convert_encoding('&#x0020;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ExclamationMark = mb_convert_encoding('&#x0021;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$QuotationMark = mb_convert_encoding('&#x0022;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$Comma = mb_convert_encoding('&#x002C;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$FullStop = mb_convert_encoding('&#x002E;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$HyphenMinus = mb_convert_encoding('&#x002D;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$Colon = mb_convert_encoding('&#x003A;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$LeftParenthesis = mb_convert_encoding('&#x005B;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$RightParenthesis = mb_convert_encoding('&#x005D;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$LeftSquareBracket = mb_convert_encoding('&#x005B;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$RightSquareBracket = mb_convert_encoding('&#x005D;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$LeftCurlyBracket = mb_convert_encoding('&#x007B;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$RightCurlyBracket = mb_convert_encoding('&#x007D;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$NoBreakSpace = mb_convert_encoding('&#x00A0;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$LeftPointingDoubleAngleQuotationMark = mb_convert_encoding('&#x00AB;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$RightPointingDoubleAngleQuotationMark = mb_convert_encoding('&#x00BB;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicComma = mb_convert_encoding('&#x060C;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterHamza = mb_convert_encoding('&#x0621;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterAlefWithMaddaAbove = mb_convert_encoding('&#x0622;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterAlefWithHamzaAbove = mb_convert_encoding('&#x0623;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterWawWithHamzaAbove = mb_convert_encoding('&#x0624;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterAlefWithHamzaBelow = mb_convert_encoding('&#x0625;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterYehWithHamzaAbove = mb_convert_encoding('&#x0626;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterAlef = mb_convert_encoding('&#x0627;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterBeh = mb_convert_encoding('&#x0628;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterTehMarbuta = mb_convert_encoding('&#x0629;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterTeh = mb_convert_encoding('&#x062A;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterTheh = mb_convert_encoding('&#x062B;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterJeem = mb_convert_encoding('&#x062C;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterHah = mb_convert_encoding('&#x062D;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterKhah = mb_convert_encoding('&#x062E;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterDal = mb_convert_encoding('&#x062F;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterThal = mb_convert_encoding('&#x0630;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterReh = mb_convert_encoding('&#x0631;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterZain = mb_convert_encoding('&#x0632;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterSeen = mb_convert_encoding('&#x0633;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterSheen = mb_convert_encoding('&#x0634;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterSad = mb_convert_encoding('&#x0635;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterDad = mb_convert_encoding('&#x0636;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterTah = mb_convert_encoding('&#x0637;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterZah = mb_convert_encoding('&#x0638;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterAin = mb_convert_encoding('&#x0639;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterGhain = mb_convert_encoding('&#x063A;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicTatweel = mb_convert_encoding('&#x0640;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterFeh = mb_convert_encoding('&#x0641;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterQaf = mb_convert_encoding('&#x0642;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterKaf = mb_convert_encoding('&#x0643;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterLam = mb_convert_encoding('&#x0644;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterMeem = mb_convert_encoding('&#x0645;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterNoon = mb_convert_encoding('&#x0646;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterHeh = mb_convert_encoding('&#x0647;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterWaw = mb_convert_encoding('&#x0648;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterAlefMaksura = mb_convert_encoding('&#x0649;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterYeh = mb_convert_encoding('&#x064A;', 'UTF-8', 'HTML-ENTITIES');

        ArabicData::$ArabicFathatan = mb_convert_encoding('&#x064B;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicDammatan = mb_convert_encoding('&#x064C;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicKasratan = mb_convert_encoding('&#x064D;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicFatha = mb_convert_encoding('&#x064E;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicDamma = mb_convert_encoding('&#x064F;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicKasra = mb_convert_encoding('&#x0650;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicShadda = mb_convert_encoding('&#x0651;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSukun = mb_convert_encoding('&#x0652;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicMaddahAbove = mb_convert_encoding('&#x0653;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicHamzaAbove = mb_convert_encoding('&#x0654;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicHamzaBelow = mb_convert_encoding('&#x0655;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicVowelSignDotBelow = mb_convert_encoding('&#x065C;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$Bullet = mb_convert_encoding('&#x2022;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterSuperscriptAlef = mb_convert_encoding('&#x0670;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterAlefWasla = mb_convert_encoding('&#x0671;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallHighLigatureSadWithLamWithAlefMaksura = mb_convert_encoding('&#x06D6;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallHighLigatureQafWithLamWithAlefMaksura = mb_convert_encoding('&#x06D7;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallHighMeemInitialForm = mb_convert_encoding('&#x06D8;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallHighLamAlef = mb_convert_encoding('&#x06D9;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallHighJeem = mb_convert_encoding('&#x06DA;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallHighThreeDots = mb_convert_encoding('&#x06DB;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallHighSeen = mb_convert_encoding('&#x06DC;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicEndOfAyah = mb_convert_encoding('&#x06DD;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicStartOfRubElHizb = mb_convert_encoding('&#x06DE;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallHighRoundedZero = mb_convert_encoding('&#x06DF;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallHighUprightRectangularZero = mb_convert_encoding('&#x06E0;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallHighMeemIsolatedForm = mb_convert_encoding('&#x06E2;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallLowSeen = mb_convert_encoding('&#x06E3;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallWaw = mb_convert_encoding('&#x06E5;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallYeh = mb_convert_encoding('&#x06E6;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallHighNoon = mb_convert_encoding('&#x06E8;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicPlaceOfSajdah = mb_convert_encoding('&#x06E9;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicEmptyCentreLowStop = mb_convert_encoding('&#x06EA;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicEmptyCentreHighStop = mb_convert_encoding('&#x06EB;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicRoundedHighStopWithFilledCentre = mb_convert_encoding('&#x06EC;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSmallLowMeem = mb_convert_encoding('&#x06ED;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicSemicolon = mb_convert_encoding('&#x061B;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterMark = mb_convert_encoding('&#x061C;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicQuestionMark = mb_convert_encoding('&#x061F;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterPeh = mb_convert_encoding('&#x067E;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterTcheh = mb_convert_encoding('&#x0686;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterVeh = mb_convert_encoding('&#x06A4;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterGaf = mb_convert_encoding('&#x06AF;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ArabicLetterNoonGhunna = mb_convert_encoding('&#x06BA;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ZeroWidthSpace = mb_convert_encoding('&#x200B;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ZeroWidthNonJoiner = mb_convert_encoding('&#x200C;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$ZeroWidthJoiner = mb_convert_encoding('&#x200D;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$LeftToRightMark = mb_convert_encoding('&#x200E;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$RightToLeftMark = mb_convert_encoding('&#x200F;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$PopDirectionalFormatting = mb_convert_encoding('&#x202C;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$LeftToRightOverride = mb_convert_encoding('&#x202D;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$RightToLeftOverride = mb_convert_encoding('&#x202E;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$NarrowNoBreakSpace = mb_convert_encoding('&#x202F;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$DottedCircle = mb_convert_encoding('&#x25CC;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$OrnateLeftParenthesis = mb_convert_encoding('&#xFD3E;', 'UTF-8', 'HTML-ENTITIES');
        ArabicData::$OrnateRightParenthesis = mb_convert_encoding('&#xFD3F;', 'UTF-8', 'HTML-ENTITIES');
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
        return implode(array_map(function($ch) { return (strlen($ch) == 3 || (hexdec(bin2hex($ch)) >= 0xD800 && hexdec(bin2hex($ch)) < 0xE000)) ? $ch : "\\x{" . bin2hex($ch) . "}"; }, preg_split('/(?<!^)(?!$)/u', $input)));
    }
    public static function MakeRegMultiEx($input)
	{
        return implode("|", $input);
    }
};
ArabicData::init();
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
        for ($count = 0; $count < count(CachedData::IslamData()->arabicnumbers->children()); $count++) {
            if (CachedData::IslamData()->arabicnumbers->children()[$count]->attributes()["name"] == $name) {
                return explode(" ", CachedData::IslamData()->arabicnumbers->children()[$count]->attributes()["text"]);
            }
        }
        return [];
    }
    public static function GetPattern($name)
    {
        for ($count = 0; $count < count(CachedData::IslamData()->arabicpatterns->children()); $count++) {
            if (CachedData::IslamData()->arabicpatterns->children()[$count]->attributes()["name"] == $name) {
                return CachedData::TranslateRegEx(CachedData::IslamData()->arabicpatterns->children()[$count]->attributes()["match"], true);
            }
        }
        return "";
    }
    public static function GetGroup($name)
    {
        for ($count = 0; $count < count(CachedData::IslamData()->arabicgroups->children()); $count++) {
            if (CachedData::IslamData()->arabicgroups->children()[$count]->attributes()["name"] == $name) {
                return array_map(function($str) use($name) { return CachedData::TranslateRegEx($str, $name == "ArabicSpecialLetters"); }, explode(" ", CachedData::IslamData()->arabicgroups->children()[$count]->attributes()["text"]));
            }
        }
        return [];
    }
    public static function GetRuleSet($name)
    {
        for ($count = 0; $count < count(CachedData::IslamData()->translitrules->children()); $count++) {
            if (CachedData::IslamData()->translitrules->children()[$count]->attributes()["name"] == $name) {
                $ruleSet = CachedData::IslamData()->translitrules->children()[$count]->children();
                for ($subCount = 0; $subCount < count($ruleSet); $subCount++) {
                    $ruleSet[$subCount]->attributes()["match"] = CachedData::TranslateRegEx($ruleSet[$subCount]->attributes()["match"], true);
                    $ruleSet[$subCount]->attributes()["evaluator"] = CachedData::TranslateRegEx($ruleSet[$subCount]->attributes()["evaluator"], false);
                    if ((string)$ruleSet[$subCount]->attributes()["rulefunc"] == "eUpperCase") {
                    	$ruleSet[$subCount]->attributes()["rulefunc"] = RuleFuncs::eUpperCase;
                    } elseif ((string)$ruleSet[$subCount]->attributes()["rulefunc"] == "eSpellNumber") {
                    	$ruleSet[$subCount]->attributes()["rulefunc"] = RuleFuncs::eSpellNumber;
                    } elseif ((string)$ruleSet[$subCount]->attributes()["rulefunc"] == "eSpellLetter") {
                    	$ruleSet[$subCount]->attributes()["rulefunc"] = RuleFuncs::eSpellLetter;
                    } elseif ((string)$ruleSet[$subCount]->attributes()["rulefunc"] == "eLookupLetter") {
                    	$ruleSet[$subCount]->attributes()["rulefunc"] = RuleFuncs::eLookupLetter;
                    } elseif ((string)$ruleSet[$subCount]->attributes()["rulefunc"] == "eLookupLongVowel") {
                    	$ruleSet[$subCount]->attributes()["rulefunc"] = RuleFuncs::eLookupLongVowel;
                    } elseif ((string)$ruleSet[$subCount]->attributes()["rulefunc"] == "eDivideTanween") {
                    	$ruleSet[$subCount]->attributes()["rulefunc"] = RuleFuncs::eDivideTanween;
                    } elseif ((string)$ruleSet[$subCount]->attributes()["rulefunc"] == "eDivideLetterSymbol") {
                    	$ruleSet[$subCount]->attributes()["rulefunc"] = RuleFuncs::eDivideLetterSymbol;
                    } elseif ((string)$ruleSet[$subCount]->attributes()["rulefunc"] == "eStopOption") {
                    	$ruleSet[$subCount]->attributes()["rulefunc"] = RuleFuncs::eStopOption;
                    } elseif ((string)$ruleSet[$subCount]->attributes()["rulefunc"] == "eLeadingGutteral") {
                    	$ruleSet[$subCount]->attributes()["rulefunc"] = RuleFuncs::eLeadingGutteral;
                    } elseif ((string)$ruleSet[$subCount]->attributes()["rulefunc"] == "eTrailingGutteral") {
                    	$ruleSet[$subCount]->attributes()["rulefunc"] = RuleFuncs::eTrailingGutteral;
                    } else {
                    	$ruleSet[$subCount]->attributes()["rulefunc"] = RuleFuncs::eNone;
                    }
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
        return preg_replace_callback("/\\{(.*?)\\}/u", function($matches) use($bAll) {
            if ($bAll) {
                if ($matches[1] == "CertainStopPattern") { return CachedData::CertainStopPattern(); }
                if ($matches[1] == "OptionalStopPattern") { return CachedData::OptionalStopPattern(); }
                if ($matches[1] == "OptionalStopPatternNotEndOfAyah") { return CachedData::OptionalStopPatternNotEndOfAyah(); }
                if ($matches[1] == "CertainNotStopPattern") { return CachedData::CertainNotStopPattern(); }
                if ($matches[1] == "OptionalNotStopPattern") { return CachedData::OptionalNotStopPattern(); }
                if ($matches[1] == "TehMarbutaStopRule") { return CachedData::TehMarbutaStopRule(); }
                if ($matches[1] == "TehMarbutaContinueRule") { return CachedData::TehMarbutaContinueRule(); }

                if ($matches[1] == "ArabicUniqueLetters") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return str_replace(ArabicData::$ArabicMaddahAbove, "", ArabicData::TransliterateFromBuckwalter($str)); }, CachedData::ArabicUniqueLetters())); }
                if ($matches[1] == "ArabicNumbers") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::TransliterateFromBuckwalter($str); }, CachedData::ArabicNumbers())); }
                if ($matches[1] == "ArabicWaslKasraExceptions") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::TransliterateFromBuckwalter($str); }, CachedData::ArabicWaslKasraExceptions())); }
                //if ($matches[1] == "SimpleSuperscriptAlefBefore") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::TransliterateFromBuckwalter(str_replace([".", """", "@", "[", "]", "-", "^"], "", $str)); }, CachedData::SimpleSuperscriptAlefBefore())); }
                //if ($matches[1] == "SimpleSuperscriptAlefNotBefore") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::TransliterateFromBuckwalter(str_replace([".", """", "@", "[", "]", "-", "^"], "", $str)); }, CachedData::SimpleSuperscriptAlefNotBefore())); }
                //if ($matches[1] == "SimpleSuperscriptAlefAfter") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::TransliterateFromBuckwalter(str_replace([".", """", "@", "[", "]", "-", "^"], "", $str)); }, CachedData::SimpleSuperscriptAlefAfter())); }
                //if ($matches[1] == "SimpleSuperscriptAlefNotAfter") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::TransliterateFromBuckwalter(str_replace([".", """", "@", "[", "]", "-", "^"], "", $str)); }, CachedData::SimpleSuperscriptAlefNotAfter())); }
                if ($matches[1] == "ArabicLongShortVowels") { return ArabicData::MakeRegMultiEx(array_map(function($strV) { return ArabicData::MakeUniRegEx($strV[0] . "(?=" . ArabicData::MakeUniRegEx($strV[1]) . ")"); }, CachedData::ArabicLongVowels())); }
                if ($matches[1] == "ArabicTanweens") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicTanweens())); }
                if ($matches[1] == "ArabicFathaDammaKasra") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicFathaDammaKasra())); }
                if ($matches[1] == "ArabicStopLetters") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicStopLetters())); }
                if ($matches[1] == "ArabicSpecialGutteral") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicSpecialGutteral())); }
                if ($matches[1] == "ArabicSpecialLeadingGutteral") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicSpecialLeadingGutteral())); }
                if ($matches[1] == "ArabicPunctuationSymbols") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicPunctuationSymbols())); }
                if ($matches[1] == "ArabicLetters") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicLetters())); }
                if ($matches[1] == "ArabicSunLettersNoLam") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicSunLettersNoLam())); }
                if ($matches[1] == "ArabicSunLetters") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicSunLetters())); }
                if ($matches[1] == "ArabicMoonLettersNoVowels") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicMoonLettersNoVowels())); }
                if ($matches[1] == "ArabicMoonLetters") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicMoonLetters())); }
                if ($matches[1] == "RecitationCombiningSymbols") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::RecitationCombiningSymbols())); }
                if ($matches[1] == "RecitationConnectingFollowerSymbols") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::RecitationConnectingFollowerSymbols())); }
                if ($matches[1] == "PunctuationSymbols") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::PunctuationSymbols())); }
                if ($matches[1] == "RecitationLettersDiacritics") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::RecitationLettersDiacritics())); }
                if ($matches[1] == "RecitationSpecialSymbolsNotStop") { return ArabicData::MakeRegMultiEx(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::RecitationSpecialSymbolsNotStop())); }
            }
            if (preg_match("/0x([0-9a-fA-F]{4})/u", $matches[1]) === 1) {
                return $bAll ? ArabicData::MakeUniRegEx(mb_convert_encoding('&#x' . substr($matches[1], 2) . ';', 'UTF-8', 'HTML-ENTITIES')) : mb_convert_encoding('&#x' . substr($matches[1], 2) . ';', 'UTF-8', 'HTML-ENTITIES');
            }
            for ($count = 0; $count < count(ArabicData::Data()->arabicletters->children()); $count++) {
                if ($matches[1] == ArabicData::Data()->arabicletters->children()[$count]->attributes()["uname"]) { return $bAll ? ArabicData::MakeUniRegEx(mb_convert_encoding('&#x' . ArabicData::Data()->arabicletters->children()[$count]->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES')) : mb_convert_encoding('&#x' . ArabicData::Data()->arabicletters->children()[$count]->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'); }
            }
           	//{0} ignore
            return $matches[0];
	    }, $value);
    }
    public static function RulesOfRecitationRegEx()
	{
        if (CachedData::$_RulesOfRecitationRegEx == null) {
            CachedData::$_RulesOfRecitationRegEx = CachedData::IslamData()->metaruleset->children();
            for ($subCount = 0; $subCount < count(CachedData::$_RulesOfRecitationRegEx); $subCount++) {
                CachedData::$_RulesOfRecitationRegEx[$subCount]->attributes()["match"] = CachedData::TranslateRegEx(CachedData::$_RulesOfRecitationRegEx[$subCount]->attributes()["match"], true);
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
	public static $_XMLDocInfo = null;
	public static function XMLDocInfo() { if (CachedData::$_XMLDocInfo === null) CachedData::$_XMLDocInfo = simplexml_load_file(dirname(__FILE__) . "/metadata/quran-data.xml"); return CachedData::$_XMLDocInfo; }
	public static $_XMLDocMain = null;
	public static function XMLDocMain() { if (CachedData::$_XMLDocMain === null) CachedData::$_XMLDocMain = simplexml_load_file(dirname(__FILE__) . "/metadata/" . TanzilReader::$QuranTextNames[0] . ".xml"); return CachedData::$_XMLDocMain; }
};

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
	const eTag = 8;
};
class RenderText
{
	public $displayClass;
	public $clr;
	public $Text;
	public $font;
	function __construct($_displayClass, $_text)
	{
		$this->displayClass = $_displayClass;
		$this->Text = $_text;
		$this->clr = "Black";
		$this->font = "";
	}
};
class RenderItem
{
	public $type;
	public $textitems;
	function __construct($_type, $_textitems)
	{
		$this->type = $_type;
		$this->textitems = $_textitems;
	}
};
class RenderArray
{
	public $Items = array();
	public static function DoRender($items)
	{
		$text = "";
		for ($count = 0; $count < count($items); $count++) {
			$text .= "<div style=\"direction: rtl; display: inline-block; padding-left: 5px; padding-right: 5px;\">";
			for ($index = 0; $index < count($items[$count]->textitems); $index++) {
				if ($items[$count]->textitems[$index]->displayClass == RenderDisplayClass::eNested) {
					$text .= RenderArray::DoRender($items[$count]->textitems[$index]->Text);
				} else {
					if (array_search($items[$count]->textitems[$index]->font, Utility::$FontList) !== false) {
						$text .= "<span dir=\"rtl\" style=\"display: block;\"><img src=\"" . plugins_url() . "" . "/islamsource/IslamSourceWP.php?Size=100&Char=" . bin2hex(mb_convert_encoding($items[$count]->textitems[$index]->Text, 'UCS-4BE', 'UTF-8')) . "&Font=" . $items[$count]->textitems[$index]->font . "\">";
					} elseif ($items[$count]->textitems[$index]->displayClass == RenderDisplayClass::eArabic) {
						$text .= "<span dir=\"rtl\" style=\"display: block; font-size: 32px;\">";
					} elseif ($items[$count]->textitems[$index]->displayClass == RenderDisplayClass::eTransliteration) {
						$text .= "<span dir=\"ltr\" style=\"display: block; font-size: 13px;\">";
					} else {
						$text .= "<span dir=\"" . ($items[$count]->textitems[$index]->displayClass == RenderDisplayClass::eLTR ? "ltr" : "rtl") . "\" style=\"display: block; font-size: 13px;\">";
					}
					if (array_search($items[$count]->textitems[$index]->font, Utility::$FontList) === false) $text .= $items[$count]->textitems[$index]->Text;
					$text .= "</span>";
				}
			}
			$text .= "</div>";
		}
		return $text;
	}
};
class RuleMetadata
{
    function __construct($newIndex, $newLength, $newType)
    {
        $this->index = $newIndex;
        $this->length = $newLength;
        $this->type = $newType;
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
class Arabic
{
	public static function TransliterateToScheme($arabicString, $schemeType, $scheme)
	{
		if ($schemeType == TranslitScheme::RuleBased) {
			return Arabic::TransliterateWithRules($arabicString, $scheme, null);
		} elseif ($schemeType == TranslitScheme::Literal) {
			return Arabic::TransliterateToRoman($arabicString, $scheme);
		} else {
			return implode(array_filter(str_split($arabicString), function($check) { return $check == " "; }));
		}
	}
    public static function GetSchemeSpecialValue($index, $scheme)
    {
        $sch = null;
        for ($count = 0; $count < count(CachedData::IslamData()->translitschemes->children()); $count++) {
            if (CachedData::IslamData()->translitschemes->children()[$count]->Name == $scheme) {
                $sch = CachedData::IslamData()->translitschemes->children()[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData()->translitschemes->children())) { return ""; }
        return explode("|", $sch->attributes()["tehmarbutaalefmaksuradaggeralefgunnah"])[$index];
    }
    public static function GetSchemeSpecialFromMatch($str, $scheme, $bExp)
    {
    	$sch = null;
        for ($count = 0; $count < count(CachedData::IslamData()->translitschemes->children()); $count++) {
            if (CachedData::IslamData()->translitschemes->children()[$count]->attributes()["name"] == $scheme) {
                $sch = CachedData::IslamData()->translitschemes->children()[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData()->translitschemes->children())) { return -1; }
        if ($bExp) {
            for ($count = 0; $count < count(CachedData::ArabicSpecialLetters()); $count++) {
                if (preg_match("/" . CachedData::ArabicSpecialLetters()[$count] . "/u", $str) == 1) { return $count; }
            }
        } else {
            if (array_search($str, CachedData::ArabicSpecialLetters()) !== false) {
                return array_search($str, CachedData::ArabicSpecialLetters());
            }
        }
        return -1;
    }
    public static function GetSchemeLongVowelFromString($str, $scheme)
    {
        $sch = null;
        for ($count = 0; $count < count(CachedData::IslamData()->translitschemes->children()); $count++) {
            if (CachedData::IslamData()->translitschemes->children()[$count]->attributes()["name"] == $scheme) {
                $sch = CachedData::IslamData()->translitschemes->children()[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData()->translitschemes->children())) { return ""; }
        if (array_search($str, CachedData::ArabicVowels()) !== false) {
            return explode("|", $sch->attributes()["fathadammakasratanweenlongvowelsdipthongsshaddasukun"])[array_search($str, CachedData::ArabicVowels())];
        }
        return "";
    }
    public static function GetSchemeGutteralFromString($str, $scheme, $leading)
    {
        $sch = null;
        for ($count = 0; $count < count(CachedData::IslamData()->translitschemes->children()); $count++) {
            if (CachedData::IslamData()->translitschemes->children()[$count]->attributes()["name"] == $scheme) {
                $sch = CachedData::IslamData()->translitschemes->children()[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData()->translitschemes->children())) { return ""; }
        if (array_search($str, CachedData::ArabicLeadingGutterals()) !== false) {
            return explode("|", $sch->attributes()["fathadammakasratanweenlongvowelsdipthongsshaddasukun"])[array_search($str, CachedData::ArabicLeadingGutterals()) + count(CachedData::ArabicVowels()) + ($leading ? count(CachedData::ArabicLeadingGutterals()) : 0)];
        }
        return "";
    }
    public static function GetSchemeValueFromSymbol($symbol, $scheme)
    {
        $sch = null;
        for ($count = 0; $count < count(CachedData::IslamData()->translitschemes->children()); $count++) {
            if (CachedData::IslamData()->translitschemes->children()[$count]->attributes()["name"] == $scheme) {
                $sch = CachedData::IslamData()->translitschemes->children()[$count];
                break;
            }
        }
        if ($count == count(CachedData::IslamData()->translitschemes->children())) { return ""; }
        if (array_search(mb_convert_encoding('&#x' . $symbol->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::ArabicLettersInOrder()) !== false) {
            return explode("|", $sch->attributes()["alphabet"])[array_search(mb_convert_encoding('&#x' . $symbol->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::ArabicLettersInOrder())];
        } elseif (array_search(mb_convert_encoding('&#x' . $symbol->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::ArabicHamzas()) !== false) {
            return explode("|", $sch->attributes()["hamza"])[array_search(mb_convert_encoding('&#x' . $symbol->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::ArabicHamzas())];
        } elseif (array_search(mb_convert_encoding('&#x' . $symbol->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::ArabicVowels()) !== false) {
            return explode("|", $sch->attributes()["fathadammakasratanweenlongvowelsdipthongsshaddasukun"])[array_search(mb_convert_encoding('&#x' . $symbol->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::ArabicVowels())];
        } elseif (array_search(mb_convert_encoding('&#x' . $symbol->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::ArabicTajweed()) !== false) {
            return explode("|", $sch->attributes()["tajweed"])[array_search(mb_convert_encoding('&#x' . $symbol->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::ArabicTajweed())];
        } elseif (array_search(mb_convert_encoding('&#x' . $symbol->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::ArabicPunctuation()) !== false) {
            return explode("|", $sch->attributes()["punctuation"])[array_search(mb_convert_encoding('&#x' . $symbol->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::ArabicPunctuation())];
        } elseif (array_search(mb_convert_encoding('&#x' . $symbol->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::NonArabicLetters()) !== false) {
            return explode("|", $sch->attributes()["nonarabic"])[array_search(mb_convert_encoding('&#x' . $symbol->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::NonArabicLetters())];
        }
        return "";
    }
	public static function TransliterateToRoman($arabicString, $scheme)
	{
        $romanString = "";
        $letters = array();
        for ($count = 0; $count < count(ArabicData::Data()->arabicletters->children()); $count++) {
        	$letters[] = ArabicData::Data()->arabicletters->children()[$count];
        }
        usort($letters, function($x, $y) use($scheme) { $compare = strlen(Arabic::GetSchemeValueFromSymbol($x, $scheme)) - strlen(Arabic::GetSchemeValueFromSymbol($y, $scheme)); if ($compare == 0) { $compare = strcmp(Arabic::GetSchemeValueFromSymbol($x, $scheme), Arabic::GetSchemeValueFromSymbol($y, $scheme)); } return $compare; });
        for ($count = 0; $count < mb_strlen($arabicString); $count++) {
            if (mb_substr($arabicString, $count, 1) == "\\") {
                $count += 1;
                if (mb_substr($arabicString, $count, 1) == ",") {
                    $romanString .= ArabicData::$ArabicComma;
                } elseif (mb_substr($arabicString, $count, 1) == ";") {
                    $romanString .= ArabicData::$ArabicSemicolon;
                } elseif (mb_substr($arabicString, $count, 1) == "?") {
                    $romanString .= ArabicData::$ArabicQuestionMark;
                } else {
                    $romanString .= mb_substr($arabicString, $count, 1);
                }
            } else {
                for ($index = 0; $index < count($letters); $index++) {
                    if (mb_substr($arabicString, $count, 1) == mb_convert_encoding('&#x' . (string)$letters[$index]->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES')) {
                        $romanString .= ($scheme == "") ? mb_convert_encoding('&#' . $letters[$index]->attributes()["extendedbuckwalter"] . ';', 'UTF-8', 'HTML-ENTITIES') : Arabic::GetSchemeValueFromSymbol($letters[$index], $scheme);
                        break;
                    }
                }
                if ($index == count($letters)) {
                    $romanString .= mb_substr($arabicString, $count, 1);
                }
            }
        }
        return $romanString;
	}
    public static function RuleFunctions()
    {
    	return [
		    function($str, $scheme) { return [strtoupper($str)]; },
		    function($str, $scheme) { return [Arabic::TransliterateWithRules(ArabicData::TransliterateFromBuckwalter(Arabic::ArabicWordFromNumber((int)(Arabic::TransliterateToScheme($str, TranslitScheme::Literal, "")), true, false, false)), $scheme, null)]; },
		    function($str, $scheme) { return [Arabic::TransliterateWithRules(Arabic::ArabicLetterSpelling($str, true), $scheme, null)]; },
		    function($str, $scheme) { return [Arabic::GetSchemeValueFromSymbol(ArabicData::Data()->arabicletters->children()[ArabicData::FindLetterBySymbol(mb_substr($str, 0, 1))], $scheme)]; },
		    function($str, $scheme) { return [Arabic::GetSchemeLongVowelFromString($str, $scheme)]; },
		    function($str, $scheme) { return [CachedData::ArabicFathaDammaKasra()[array_search($str, CachedData::ArabicTanweens)], ArabicData::$ArabicLetterNoon]; },
		    function($str, $scheme) { return ["", ""]; },
		    function($str, $scheme) { return [""]; },
		    function($str, $scheme) { return [Arabic::GetSchemeGutteralFromString(mb_substr($str, 0, mb_strlen($str) - 1), $scheme, true) . mb_substr($str, mb_strlen($str) - 1, 1)]; },
		    function($str, $scheme) { return [mb_substr($str, 0, 1) . Arabic::GetSchemeGutteralFromString(mb_substr($str, 1), $scheme, false)]; }
			];
    }
    public static $AllowZeroLength = ["helperlparen", "helperrparen"];
	public static function IsLetter($index)
	{
        return array_search(mb_convert_encoding('&#x' . ArabicData::Data()->arabicletters->children[$index]->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::ArabicLetters()) !== false;
    }
    public static function IsPunctuation($index)
    {
        return array_search(mb_convert_encoding('&#x' . ArabicData::Data()->arabicletters->children[$index]->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::PunctuationSymbols()) !== false;
    }
    public static function IsStop($index)
    {
        return array_search(mb_convert_encoding('&#x' . ArabicData::Data()->arabicletters->children[$index]->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::ArabicStopLetters()) !== false;
    }
    public static function IsWhitespace($index)
    {
        return array_search(mb_convert_encoding('&#x' . ArabicData::Data()->arabicletters->children[$index]->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES'), CachedData::WhitespaceSymbols()) !== false;
    }

    public static function ArabicLetterSpelling($input, $quranic)
    {
        $output = "";
        foreach ($input as $ch) {
            $index = ArabicData::FindLetterBySymbol($ch);
            if ($index != -1 && ArabicData::IsLetter($index)) {
                if ($output != "" && !$quranic) { $output .= " "; }
                $output .= $quranic ? substr(ArabicData::Data()->arabicletters->children()[$index]->attributes()["symbolname"], 0, strlen(ArabicData::Data()->arabicletters->children()[$index]->attributes()["symbolname"]) - 1) . ((substr(ArabicData::Data()->arabicletters->children()[$index]->attributes()["symbolname"], -1) == "n") ? "" : "o") : ArabicData::Data()->arabicletters->children()[$index]->attributes()["symbolname"];
            } elseif ($index != -1 && mb_convert_encoding('&#x' . ArabicData::Data()->arabicletters->children()[$index]->attributes()["symbol"] . ';', 'UTF-8', 'HTML-ENTITIES') == ArabicData::$ArabicMaddahAbove) {
                if (!$quranic) { $output .= $ch; }
            }
        }
        return ArabicData::TransliterateFromBuckwalter($output);
    }
    public static function ArabicWordForLessThanThousand($number, $useClassic, $useAlefHundred)
    {
        $str = "";
        $hundStr = "";
        if ($number >= 100) {
            $hundStr = $useAlefHundred ? substr_replace(CachedData::ArabicBaseHundredNumbers()[(int)($number / 100) - 1], "A", 2, 0) : CachedData::ArabicBaseHundredNumbers()[(int)($number / 100) - 1];
            if (($number % 100) == 0) { return $hundStr; }
            $number = $number % 100;
        }
        if (($number % 10) != 0 && $number != 11 && $number != 12) {
            $str = CachedData::ArabicBaseNumbers()[($number % 10) - 1];
        }
        if ($number >= 11 && $number < 20) {
            if ($number == 11 || $number == 12) {
                $str .= CachedData::ArabicBaseExtraNumbers()[$number - 11];
            } else {
                $str = substr($str, 0, strlen($str) - 1) . "a";
            }
            $str .= " Ea\$ara";
        } elseif (($number == 0 && $str == "") || $number == 10 || $number >= 20) {
            $str = ($str == "" ? "" : $str . " wa") . CachedData::ArabicBaseTenNumbers()[(int)($number / 10)];
        }
        return $useClassic ? ($str == "" ? "" : $str . ($hundStr == "" ? "" : " wa")) . $hundStr : ($hundStr == "" ? "" : $hundStr . ($str == "" ? "" : " wa")) . $str;
    }
    public static function ArabicWordFromNumber($number, $useClassic, $useAlefHundred, $useMilliard)
    {
        $str = "";
        $nextStr = "";
        $curBase = 3;
        $baseNums = [1000, 1000000, 1000000000, 1000000000000];
        $bases = [CachedData::ArabicBaseThousandNumbers(), CachedData::ArabicBaseMillionNumbers(), $useMilliard ? CachedData::ArabicBaseMilliardNumbers() : CachedData::ArabicBaseBillionNumbers(), CachedData::ArabicBaseTrillionNumbers()];
        do {
            if ($number >= $baseNums[$curBase] && $number < 2 * $baseNums[$curBase]) {
                $nextStr = $bases[$curBase][0];
            } elseif ($number >= 2 * $baseNums[$curBase] && $number < 3 * $baseNums[$curBase]) {
                $nextStr = $bases[$curBase][1];
            } elseif ($number >= 3 * $baseNums[$curBase] && $number < 10 * $baseNums[$curBase]) {
                $nextStr = substr(CachedData::ArabicBaseNumbers((int)($number / $baseNums[$curBase] - 1)), 0, strlen(CachedData::ArabicBaseNumbers()[(int)($number / $baseNums[$curBase] - 1)]) - 1) . "u " . $bases[$curBase][2];
            } elseif ($number >= 10 * $baseNums[$curBase] && $number < 11 * $baseNums[$curBase]) {
                $nextStr = substr(CachedData::ArabicBaseTenNumbers()[1], 0, strlen(CachedData::ArabicBaseTenNumbers()[1]) - 1) . "u " . $bases[$curBase][2];
            } elseif ($number >= $baseNums[$curBase]) {
                $nextStr = Arabic::ArabicWordForLessThanThousand((int)((int)($number / $baseNums[$curBase]) % 100), $useClassic, $useAlefHundred);
                if ($number >= 100 * $baseNums[$curBase] && $number < ($useClassic ? 200 : 101) * $baseNums[$curBase]) {
                    $nextStr = substr($nextStr, 0, strlen($nextStr) - 1) . "u " . substr($bases[$curBase][0], 0, strlen($bases[$curBase][0]) - 1) . "K";
                } elseif ($number >= 200 * $baseNums[$curBase] && $number < ($useClassic ? 300 : 201) * $baseNums[$curBase]) {
                    $nextStr = substr($nextStr, 0, strlen($nextStr) - 2) . " " . substr($bases[$curBase][0], 0, strlen($bases[$curBase][0]) - 1) . "K";
                } elseif ($number >= 300 * $baseNums[$curBase] && ($useClassic || (int)($number / $baseNums[$curBase]) % 100 == 0)) {
                    $nextStr = substr($nextStr, 0, strlen($nextStr) - 1) . "i " . substr($bases[$curBase][0], 0, strlen($bases[$curBase][0]) - 1) . "K";
                } else {
                    $nextStr .= " " . substr($bases[$curBase][0], 0, strlen($bases[$curBase][0]) - 1) . "FA";
                }
            }
            $number = $number % $baseNums[$curBase];
            $curBase -= 1;
            $str = $useClassic ? ($nextStr == "" ? "" : $nextStr . ($str == "" ? "" : " wa")) . $str : ($str == "" ? "" : $str . ($nextStr == "" ? "" : " wa")) . $nextStr;
            $nextStr = "";
        } while ($curBase >= 0);
        if ($number != 0 || $str == "") { $nextStr = Arabic::ArabicWordForLessThanThousand((int)($number), $useClassic, $useAlefHundred); }
        return $useClassic ? ($nextStr == "" ? "" : $nextStr . ($str == "" ? "" : " wa")) . $str : ($str == "" ? "" : $str . ($nextStr == "" ? "" : " wa")) . $nextStr;
    }
    public static function ReplaceMetadata($arabicString, $metadataRule, $scheme, $optionalStops)
    {
        for ($count = 0; $count < count(CachedData::ColoringSpelledOutRules()); $count++) {
            $match = current(array_filter(explode("|", CachedData::ColoringSpelledOutRules()[$count]->attributes()["match"]), function($str) use($metadataRule) { return array_search($str, array_map(function($s) use($metadataRule) { return preg_replace("/\\(.*\\)/u", "", $s); }, explode("|", $metadataRule->type))) !== false; }));
            if ($match != null) {
                $str = sprintf(preg_replace_callback('/\\{(0|[1-9]\\d*)\\}/u', function($match) { return "%" . (((int)$match[1]) + 1) . "\$s"; }, CachedData::ColoringSpelledOutRules()[$count]->attributes()["evaluator"]), mb_substr($arabicString, $metadataRule->index, $metadataRule->length));
                if ((int)CachedData::ColoringSpelledOutRules()[$count]->attributes()["rulefunc"] != RuleFuncs::eNone) {
                    $args = Arabic::RuleFunctions()[(int)CachedData::ColoringSpelledOutRules()[$count]->attributes()["rulefunc"] - 1];
                    $args = $args($str, $scheme);
                    if (count($args) == 1) {
                        $str = $args[0];
                    } else {
                        $metaArgs = explode(",", preg_match("/" . $match . "\\((.*)\\)" . "/u", $metadataRule->type)[1]);
                        $str = "";
                        for ($index = 0; $index < count($args); $index++) {
                            if ($args[$index] != null) {
                                $str .= ArabicData::ReplaceMetadata($args[$index], new RuleMetadata(0, mb_strlen($args[$index]), str_replace(" ", "|", $metaArgs[$index])), $scheme, $optionalStops);
                            }
                        }
                    }
                }
                $arabicString = mb_substr($arabicString, 0, $metadataRule->index) . $str . mb_substr($arabicString, $metadataRule->index + $metadataRule->length);
            }
        }
        return $arabicString;
    }
    public static function MatchResult($eval, $offset, $str, $matches)
    {
		return preg_replace_callback("/\\$(\\$|&|`|\\'|[0-9]+)/u", function($m) use ($offset, $str, $matches) { if ($m[1] === '$') return '$'; if ($m[1] === '`') return mb_substr($str, 0, $offset); if ($m[1] === '\'') return mb_substr($str, $offset + mb_strlen($matches[0])); if ($m[1] === '&' || (int)($m[1]) <= 0 || (int)($m[1]) >= count($matches)) return $matches[0]; return $matches[(int)($m[1])]; }, $eval);
    }
    public static function DoErrorCheck($arabicString)
    {
        //need to check for decomposed first
        for ($count = 0; $count < count(CachedData::ErrorCheckRules()); $count++) {
            preg_match_all("/" . CachedData::ErrorCheckRules()[$count]->attributes()["match"] . "/u", $arabicString, $matches, PREG_OFFSET_CAPTURE | PREG_SET_ORDER);
            for ($matchIndex = 0; $matchIndex < count($matches); $matchIndex++) {
                if (CachedData::ErrorCheckRules()[$count]->attributes()["negativematch"] == null || Arabic::MatchResult(CachedData::ErrorCheckRules()[$count]->attributes()["negativematch"], $matches[$matchIndex][0][1], $arabicString, array_map(function($match) { return $match[0]; }, $matches[$matchIndex])) == "") {
                    //Debug.Print(ErrorCheckRules[$count]->attributes()["name"] . ": " . substr_replace(Arabic::TransliterateToScheme($arabicString, TranslitScheme::Literal, ""), "<!-- -->", $matches[$matchIndex][0][1], 0))
                }
            }
        }
    }
    public static function JoinContig($arabicString, $preString, $postString)
    {
        $index = strrpos($preString, " ");
        //take last word of pre string and first word of post string or another if it is a pause marker
        //end of ayah sign without number is used as a proper place holder
        if ($index != -1 && strlen($preString) - 2 == $index) { $index = strrpos($preString, " ", $index - 1); }
        if ($index != -1) { $preString = substr($preString, $index + 1); }
        if ($preString != "") { $preString .= " " . ArabicData::$ArabicEndOfAyah . " "; }
        $index = strpos($postString, " ");
        if ($index == 2) { $index = strpos($preString, " ", $index + 1); }
        if ($index != -1) { $postString = substr($postString, 0, $index); }
        if ($postString != "") { $postString = " " . ArabicData::$ArabicEndOfAyah . " " . $postString; }
        return $preString . $arabicString . $postString;
    }
    public static function UnjoinContig($arabicString, $preString, $postString)
    {
        $index = strpos($arabicString, ArabicData::$ArabicEndOfAyah);
        if ($preString != "" && $index != -1) {
            $arabicString = substr($arabicString, $index + 1 + 1);
        }
        $index = strrpos($arabicString, ArabicData::$ArabicEndOfAyah);
        if ($postString != "" && $index != -1) {
            $arabicString = substr($arabicString, 0, $index - 1);
        }
        return $arabicString;
    }
    public static function TransliterateContigWithRules($arabicString, $preString, $postString, $scheme, $optionalStops, $preOptionalStops, $postOptionalStops)
    {
        return Arabic::UnjoinContig(Arabic::TransliterateWithRules(Arabic::JoinContig($arabicString, $preString, $postString), $scheme, null), $preString, $postString);
    }
    public static function TransliterateWithRules($arabicString, $scheme, $optionalStops)
    {
        $metadataList = array();
        Arabic::DoErrorCheck($arabicString);
        for ($count = 0; $count < count(CachedData::RulesOfRecitationRegEx()); $count++) {
            if (CachedData::RulesOfRecitationRegEx()[$count]->attributes()["evaluator"] != null) {
                preg_match_all("/" . CachedData::RulesOfRecitationRegEx()[$count]->attributes()["match"] . "/u", $arabicString, $matches, PREG_OFFSET_CAPTURE | PREG_SET_ORDER);
                for ($matchIndex = 0; $matchIndex < count($matches); $matchIndex++) {
                    for ($subCount = 0; $subCount < count(explode(";", CachedData::RulesOfRecitationRegEx()[$count]->attributes()["evaluator"])); $subCount++) {
                        if (explode(";", CachedData::RulesOfRecitationRegEx()[$count]->attributes()["evaluator"])[$subCount] != null && (strlen($matches[$matchIndex][$subCount + 1][0]) != 0 || array_search(explode(";", CachedData::RulesOfRecitationRegEx()[$count]->attributes()["evaluator"])[$subCount], Arabic::$AllowZeroLength) !== false)) {
                            array_push($metadataList, new RuleMetadata(mb_strlen(substr($arabicString, 0, $matches[$matchIndex][$subCount + 1][1])), mb_strlen($matches[$matchIndex][$subCount + 1][0]), explode(";", CachedData::RulesOfRecitationRegEx()[$count]->attributes()["evaluator"])[$subCount]));
                        }
                    }
                }
            }
        }
        usort($metadataList, function($x, $y) {
		        if ($x->index == $y->index) {
		            return $y->length == $x->length ? 0 : ($y->length > $x->length ? 1 : -1);
		        } else {
		            return $y->index > $x->index ? 1 : -1;
		        }
		    });
        for ($index = 0; $index < count($metadataList); $index++) {
            $arabicString = Arabic::ReplaceMetadata($arabicString, $metadataList[$index], $scheme, $optionalStops);
        }
        //redundant romanization rules should have -'s such as seen/teh/kaf-heh
        for ($count = 0; $count < count(CachedData::RomanizationRules()); $count++) {
            if ((int)CachedData::RomanizationRules()[$count]->attributes()["rulefunc"] == RuleFuncs::eNone) {
                $arabicString = preg_replace("/" . CachedData::RomanizationRules()[$count]->attributes()["match"] . "/u", CachedData::RomanizationRules()[$count]->attributes()["evaluator"], $arabicString);
            } else {
                $arabicString = preg_replace_callback("/" . CachedData::RomanizationRules()[$count]->attributes()["match"] . "/u", function($matches) use($count, $scheme) { $func = Arabic::RuleFunctions()[(int)CachedData::RomanizationRules()[$count]->attributes()["rulefunc"] - 1]; return $func(Arabic::MatchResult(CachedData::RomanizationRules()[$count]->attributes()["evaluator"], mb_strpos($matches[0], CachedData::RomanizationRules()[$count]->attributes()["match"]), CachedData::RomanizationRules()[$count]->attributes()["match"], $matches), $scheme)[0]; }, $arabicString);
            }
        }

        //process wasl loanwords and names
        //process loanwords and names
        return $arabicString;
    }
    static function GetTransliterationTable($scheme)
    {
        $items = array();
        $items.AddRange(array_map(function($letter) { return new RenderItem(RenderTypes::eText, [ new RenderText(RenderDisplayClass::eArabic, Letter), new RenderText(RenderDisplayClass::eTransliteration, Arabic::GetSchemeValueFromSymbol(ArabicData::Data()->arabicletters->children()[ArabicData::FindLetterBySymbol($letter[0])], $scheme))]); }, CachedData::ArabicLettersInOrder));
        $items.AddRange(array_map(function($combo) { return new RenderItem(RenderTypes::eText, [ new RenderText(RenderDisplayClass::eArabic, preg_replace("/\\(?\\\\u([0-9a-fA-F]{4})\\)?/u", function($match) { return mb_convert_encoding('&#x' . $match[1] . ';', 'UTF-8', 'HTML-ENTITIES'); }, str_replace([CachedData::TehMarbutaStopRule, CachedData::TehMarbutaContinueRule], ["", "..."], $combo))), new RenderText(RenderDisplayClass::eTransliteration, Arabic::GetSchemeSpecialValue(Arabic::GetSchemeSpecialFromMatch($combo, $scheme, false), $scheme))]); }, CachedData::ArabicSpecialLetters));
        $items.AddRange(array_map(function($letter) { return new RenderItem(RenderTypes::eText, [ new RenderText(RenderDisplayClass::eArabic, Letter), new RenderText(RenderDisplayClass::eTransliteration, Arabic::GetSchemeValueFromSymbol(ArabicData::Data()->arabicletters->children()[ArabicData::FindLetterBySymbol($letter[0])], $scheme))]); }, CachedData::ArabicHamzas));
        $items.AddRange(array_map(function($combo) { return new RenderItem(RenderTypes::eText, [ new RenderText(RenderDisplayClass::eArabic, Combo), new RenderText(RenderDisplayClass::eTransliteration, Arabic::GetSchemeLongVowelFromString($combo, $scheme))]); }, CachedData::ArabicVowels));
        $items.AddRange(array_map(function($letter) { return new RenderItem(RenderTypes::eText, [ new RenderText(RenderDisplayClass::eArabic, Letter), new RenderText(RenderDisplayClass::eTransliteration, Arabic::GetSchemeValueFromSymbol(ArabicData::Data()->arabicletters->children()[ArabicData::FindLetterBySymbol($letter[0])], $scheme))]); }, CachedData::ArabicTajweed));
        return $items;
    }
    public static function GetTransliterationSchemes()
    {
        $strings = array();
        $strings[0] = [Utility::LoadResourceString("IslamSource_Off"), "0"];
        $strings[1] = [Utility::LoadResourceString("IslamSource_ExtendedBuckwalter"), "1"];
        for ($count = 0; $count < count(CachedData::IslamData()->translitschemes->children()); $count++) {
            $strings[$count * 2 + 2] = [Utility::LoadResourceString("IslamSource_" . CachedData::IslamData()->translitschemes->children()[$count]->attributes()["name"]), strval($count * 2 + 2)];
            $strings[$count * 2 + 1 + 2] = [Utility::LoadResourceString("IslamSource_" . CachedData::IslamData()->translitschemes->children()[$count]->attributes()["name"]) . " Literal", strval($count * 2 + 1 + 2)];
        }
        return $strings;
    }
};
class Languages
{
    public static function GetLanguageInfoByCode($code)
    {
        for ($count = 0; $count < count(CachedData::IslamData()->languages->children()); $count++) {
            if (CachedData::IslamData()->languages->children()[$count]->attributes()["code"] == $code) { return CachedData::IslamData()->languages->children()[$count]; }
        }
        return null;
    }
};
class TanzilReader
{
	public static function IsQuranTextReference($str) { return preg_match("/^(?:,?(\\d+)(?:\\:(\\d+))?(?:\\:(\\d+))?(?:-(\\d+)(?:\\:(\\d+))?(?:\\:(\\d+))?)?)+$/u", $str); }
	public static function QuranTextFromReference($str, $schemetype, $scheme, $translationindex)
	{
		$renderer = new RenderArray("");
		preg_match_all('/(?:,?(\\d+)(?:\\:(\\d+))?(?:\\:(\\d+))?(?:-(\\d+)(?:\\:(\\d+))?(?:\\:(\\d+))?)?)/u', $str, $matches, PREG_SET_ORDER);
		for ($count = 0; $count < count($matches); $count++) {
			$basechapter = $matches[$count][1];
			$baseverse = array_key_exists(2, $matches[$count]) ? $matches[$count][2] : 0;
			$wordnumber = array_key_exists(3, $matches[$count]) ? $matches[$count][3] : 0;
			$endchapter = array_key_exists(4, $matches[$count]) ? $matches[$count][4] : 0;
			$extraversenumber = array_key_exists(5, $matches[$count]) ? $matches[$count][5] : 0;
			$endwordnumber = array_key_exists(6, $matches[$count]) ? $matches[$count][6] : 0;
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
	        $renderer->Items = array_merge($renderer->Items, TanzilReader::DoGetRenderedQuranText(TanzilReader::QuranTextRangeLookup($basechapter, $baseverse, $wordnumber, $endchapter, $extraversenumber, $endwordnumber), $basechapter, $baseverse, CachedData::IslamData()->translations->children()[$translationindex]->attributes()["file"], $schemetype, $scheme, $translationindex)->Items);
            $reference = (string)($basechapter) . ($baseverse != 0 ? ":" . (string)($baseverse) : "") . ($endchapter != 0 ? "-" . (string)($endchapter) . ($extraversenumber != 0 ? ":" . (string)($extraversenumber) : "") : ($extraversenumber != 0 ? "-" . (string)($extraversenumber) : ""));
            array_push($renderer->Items, new RenderItem(RenderTypes::eHeaderCenter, [new RenderText(RenderDisplayClass::eLTR, "(Qur'an " . $reference . ")")]));
	    }
		return $renderer;
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
            $verseIndex = 0;
            for ($wordcount = 1; $wordcount <= $wordnumber - 1; $wordcount++) {
                $verseIndex = mb_strpos($qurantext[0][0], " ", $verseIndex) + 1;
            }
            $qurantext[0][0] = preg_replace("/" . implode("|", array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicStopLetters())) . "|" . ArabicData::$ArabicStartOfRubElHizb . "|" . ArabicData::$ArabicPlaceOfSajdah . "/u", "\0", preg_replace("/(^\\s*|\\s+)[^\\s" . implode(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicStopLetters())) . ArabicData::$ArabicStartOfRubElHizb . ArabicData::$ArabicPlaceOfSajdah . "]+(?=\\s*$|\\s+)/u", "$1", mb_substr($qurantext[0][0], 0, $verseIndex))) . mb_substr($qurantext[0][0], $verseIndex);
		}
		if ($endwordnumber != 0) {
            $verseIndex = 0;
            //selections are always within the same chapter
            $lastChapter = count($qurantext) - 1;
            $lastVerse = (int)($extraversenumber != 0 ? count($qurantext[$lastChapter]) - 1 : 0);
            while (hexdec(bin2hex(mb_substr($qurantext[$lastChapter][$lastVerse], $verseIndex, 1))) == 0 || mb_substr($qurantext[$lastChapter][$lastVerse], $verseIndex, 1) == " ") {
                $verseIndex += 1;
            }
            for ($wordcount = $wordnumber - 1; $wordcount <= $endwordnumber - 1; $wordcount++) {
                $verseIndex = mb_strpos($qurantext[$lastChapter][$lastVerse], " ", $verseIndex) + 1;
            }
            if ($verseIndex == 0) { $verseIndex = mb_strlen($qurantext[$lastChapter][$lastVerse]); }
            $qurantext[$lastChapter][$lastVerse] = mb_substr($qurantext[$lastChapter][$lastVerse], 0, $verseIndex) . preg_replace("/" . implode("|", array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicStopLetters())) . "|" . ArabicData::$ArabicStartOfRubElHizb . "|" . ArabicData::$ArabicPlaceOfSajdah . "/u", "\0", preg_replace("/(^\\s*|\\s+)[^\\s" . implode(array_map(function($str) { return ArabicData::MakeUniRegEx($str); }, CachedData::ArabicStopLetters())) . ArabicData::$ArabicStartOfRubElHizb . ArabicData::$ArabicPlaceOfSajdah . "]+(?=\\s*$|\\s+)/u", "$1", mb_substr($qurantext[$lastChapter][$lastVerse], $verseIndex)));
		}
		return $qurantext;
	}
	public static function GetQuranTextRange($xmldoc, $startchapter, $startverse, $endchapter, $endverse)
	{
		if ($startchapter == -1) $startchapter = 1;
		if ($endchapter == -1) $endchapter = TanzilReader::GetChapterCount();
		$chapterverses = array();
        for ($count = $startchapter; $count <= $endchapter; $count++) {
            array_push($chapterverses, TanzilReader::GetQuranText($xmldoc, $count, $startchapter == $count ? $startverse : -1, $endchapter == $count ? $endverse : -1));
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
		return CachedData::IslamData()->translations->children()[$index]->attributes()["file"] . ".txt";
	}
	public static function DoGetRenderedQuranText($qurantext, $basechapter, $baseverse, $translation, $schemetype, $scheme, $translationindex)
	{
        $renderer = new RenderArray("");
		$lines = explode("\n", file_get_contents(dirname(__FILE__) . "/metadata/" . TanzilReader::GetTranslationFileName($translation)));
		$w4wlines = explode("\n", file_get_contents(dirname(__FILE__) . "/metadata/en.w4w.shehnazshaikh.txt"));
		if ($qurantext !== null) {
			for ($chapter = 0; $chapter < count($qurantext); $chapter++) {
                $chapterNode = TanzilReader::GetChapterByIndex($basechapter + $chapter);
                array_push($renderer->Items, new RenderItem(RenderTypes::eHeaderLeft, [new RenderText(RenderDisplayClass::eArabic, ArabicData::TransliterateFromBuckwalter("'aAya`tuhaA " . $chapterNode->attributes()["ayas"] . " ")), new RenderText(RenderDisplayClass::eTransliteration, trim(Arabic::TransliterateToScheme(ArabicData::TransliterateFromBuckwalter("'aAya`tuhaA " . $chapterNode->attributes()["ayas"] . " "), $schemetype, $scheme))), new RenderText(RenderDisplayClass::eLTR, "Verses " . $chapterNode->attributes()["ayas"] . " ")]));
                array_push($renderer->Items, new RenderItem(RenderTypes::eHeaderCenter, [new RenderText(RenderDisplayClass::eArabic, ArabicData::TransliterateFromBuckwalter("suwrapu " . CachedData::IslamData()->quranchapters->children()[(int)($chapterNode->attributes()["index"]) - 1]->attributes()["name"] . " ")), new RenderText(RenderDisplayClass::eTransliteration, trim(Arabic::TransliterateToScheme(ArabicData::TransliterateFromBuckwalter("suwrapu " . CachedData::IslamData()->quranchapters->children()[(int)($chapterNode->attributes()["index"]) - 1]->attributes()["name"] . " "), $schemetype, $scheme))), new RenderText(RenderDisplayClass::eLTR, "Chapter " . TanzilReader::GetChapterEName($chapterNode) . " ")]));
                array_push($renderer->Items, new RenderItem(RenderTypes::eHeaderRight, [new RenderText(RenderDisplayClass::eArabic, ArabicData::TransliterateFromBuckwalter("rukuwEaAtuhaA " . $chapterNode->attributes()["rukus"] . " ")), new RenderText(RenderDisplayClass::eTransliteration, trim(Arabic::TransliterateToScheme(ArabicData::TransliterateFromBuckwalter("rukuwEaAtuhaA " . $chapterNode->attributes()["rukus"] . " "), $schemetype, $scheme))), new RenderText(RenderDisplayClass::eLTR, "Rukus " . $chapterNode->attributes()["rukus"] . " ")]));
				for ($verse = 0; $verse < count($qurantext[$chapter]); $verse++) {
                    $items = array();
                    $text = "";
                    //hizb symbols not needed as Quranic text already contains them
                    //if ($basechapter + $chapter != 1 && TanzilReader::IsQuarterStart($base$chapter + $chapter, (int)(($chapter == 0 ? $baseverse : 1)) + $verse) {
                    //    $text .= ArabicData::TransliterateFromBuckwalter("B");
                    //    array_push($items, new RenderItem(RenderTypes::eText, [new RenderText(RenderDisplayClass::eArabic, ArabicData::TransliterateFromBuckwalter("B"))]));
                    //}
                    if ((int)($chapter == 0 ? $baseverse : 1) + $verse == 1) {
                        $node = TanzilReader::GetTextVerse(TanzilReader::GetTextChapter(CachedData::XMLDocMain(), $basechapter + $chapter), 1)->attributes()["bismillah"];
                        if ($node != null) {
                            array_push($renderer->Items, new RenderItem(RenderTypes::eText, [new RenderText(RenderDisplayClass::eArabic, $node . " "), new RenderText(RenderDisplayClass::eTransliteration, trim(Arabic::TransliterateToScheme($node, $schemetype, $scheme))), new RenderText((TanzilReader::IsTranslationTextLTR($translationindex) ? RenderDisplayClass::eLTR : RenderDisplayClass::eRTL), TanzilReader::GetTranslationVerse($lines, 1, 1))]));
                        }
                    }
                    $words = ($qurantext[$chapter][$verse] == null ? [] : explode(" ", $qurantext[$chapter][$verse]));
                    $translitWords = explode(" ", Arabic::TransliterateToScheme($qurantext[$chapter][$verse], $schemetype, $scheme));
                    $pauseMarks = 0;
                    for ($count = 0; $count < count($words); $count++) {
                        //handle start/end words here which have space placeholders
                        if (strlen($words[$count]) == 1 && $words[$count][0] == "\0") {
                            $pauseMarks += 1;
                        } elseif (strlen($words[$count]) == 1 &&
                            (ArabicData::IsStop(ArabicData::FindLetterBySymbol($words[$count][0])) || $words[$count][0] == ArabicData::$ArabicStartOfRubElHizb || $words[$count][0] == ArabicData::$ArabicPlaceOfSajdah)) {
                            $pauseMarks += 1;
                            array_push($items, new RenderItem(RenderTypes::eText, [new RenderText(RenderDisplayClass::eArabic, " " . $words[$count]), new RenderText(RenderDisplayClass::eTransliteration, $translitWords[$count])]));
                        } elseif (strlen($words[$count]) != 0) {
                            array_push($items, new RenderItem(RenderTypes::eText, [new RenderText(RenderDisplayClass::eArabic, $words[$count]), new RenderText(RenderDisplayClass::eTransliteration, $translitWords[$count]), new RenderText((TanzilReader::IsTranslationTextLTR($translationindex) ? RenderDisplayClass::eLTR : RenderDisplayClass::eRTL), TanzilReader::GetW4WTranslationVerse($w4wlines, $basechapter + $chapter, (int)(($chapter == 0 ? $baseverse : 1)) + $verse, $count - $pauseMarks))]));
                        }
                    }
                    $text .= $qurantext[$chapter][$verse] . " ";
                    if (TanzilReader::IsSajda($basechapter + $chapter, (int)(($chapter == 0 ? $baseverse : 1)) + $verse)) {
                        //Sajda markers are already in the text
                        //Text .= ArabicData::TransliterateFromBuckwalter("R")
                        //array_push($items, new RenderItem(RenderTypes::eText, [new RenderText(RenderDisplayClass::eArabic, ArabicData::TransliterateFromBuckwalter("R"))]));
                    }
                    $text .= ArabicData::TransliterateFromBuckwalter("=" . strval((int)(($chapter == 0 ? $baseverse : 1)) + $verse)) . " ";
                    array_push($items, new RenderItem(RenderTypes::eText, [new RenderText(RenderDisplayClass::eArabic, ArabicData::TransliterateFromBuckwalter("=" . strval((int)(($chapter == 0 ? $baseverse : 1)) + $verse))), new RenderText((TanzilReader::IsTranslationTextLTR($translationindex) ? RenderDisplayClass::eLTR : RenderDisplayClass::eRTL), "(" . strval((int)($chapter == 0 ? $baseverse : 1) + $verse) . ")")]));
                    //$text .= ArabicData::TransliterateFromBuckwalter("(" . strval(($chapter == 0 ? $baseverse : 1) + $verse) . ") ")
                    array_push($renderer->Items, new RenderItem(RenderTypes::eText, [new RenderText(RenderDisplayClass::eNested, $items), new RenderText(RenderDisplayClass::eArabic, $text), new RenderText(RenderDisplayClass::eTransliteration, trim(Arabic::TransliterateToScheme($qurantext[$chapter][$verse] . " " . ArabicData::TransliterateFromBuckwalter("=" . strval((int)($chapter == 0 ? $baseverse : 1) + $verse)) . " ", $schemetype, $scheme))), new RenderText((TanzilReader::IsTranslationTextLTR($translationindex) ? RenderDisplayClass::eLTR : RenderDisplayClass::eRTL), "(" . strval((int)($chapter == 0 ? $baseverse : 1) + $verse) . ") " . TanzilReader::GetTranslationVerse($lines, $basechapter + $chapter, (int)($chapter == 0 ? $baseverse : 1) + $verse))]));
				}
			}
		}
		return $renderer;
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
		return Utility::GetChildNodeByIndex("sura", "index", $index, CachedData::XMLDocInfo()->children()->suras->children());
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
        return Utility::GetChildNodeCount("sura", CachedData::XMLDocInfo()->children()->suras->children());
    }
    public static function IsTranslationTextLTR($index)
    {
        return !Languages::GetLanguageInfoByCode(substr(CachedData::IslamData()->translations->children()[$index]->attributes()["file"], 0, 2))->IsRTL;
    }
    public static function GetChapterEName($chapterNode)
    {
        return Utility::LoadResourceString("IslamInfo_QuranChapter" . $chapterNode->attributes()["index"]);
    }
    public static function GetW4WTranslationVerse($lines, $chapter, $verse, $word)
    {
        $words = explode("|", $lines[TanzilReader::GetVerseNumber($chapter, $verse) - 1]);
        if ($word >= count($words)) {
            return "";
        } else {
            return $words[$word];
        }
    }
    public static function IsSajda($chapter, $verse)
    {
        for ($count = 0; $count < count(CachedData::XMLDocInfo()->children()->sajdas->children()); $count++) {
            $node = CachedData::XMLDocInfo()->children()->sajdas->children()[$count];
            if ($node->getName() == "sajda" &&
                (int)($node->attributes()["sura"]) == $chapter &&
                (int)($node->attributes()["aya"]) == $verse) {
                return true;
            }
        }
        return false;
    }
    public static $QuranTextNames = array("quran-hafs", "quran-warsh", "quran-alduri");
};

class DocBuilder
{
	public static function BuckwalterTextFromReferences($ID, $schemetype, $scheme, $strings, $translationID, $translationindex)
	{
        $renderer = new RenderArray($ID);
        if ($strings == null) { return $renderer; }
		if (preg_match_all('/(.*?)(?:(\\{)(.*?)(\\})|$)/u', $strings, $matches, PREG_SET_ORDER) !== 0) {
			for ($matchcount = 0; $matchcount < count($matches); $matchcount++) {
	            if (count($matches[$matchcount])) {
	                if (array_key_exists(1, $matches[$matchcount])) {
	                    $englishByWord = explode("|", Utility::LoadResourceString("IslamInfo_" . $translationID . "WordByWord"));
	                    $arabicText = explode(" ", $matches[$matchcount][1]);
	                    $transliteration = explode(" ", Arabic::TransliterateToScheme(ArabicData::TransliterateFromBuckwalter($matches[$matchcount][1]), $schemetype, $scheme));
	                    array_push($renderer->Items, new RenderItem(RenderTypes::eHeaderCenter, [new RenderText(RenderDisplayClass::eLTR, Utility::LoadResourceString("IslamInfo_" . $translationID))]));
	                    $items = array();
	                    for ($wordcount = 0; $wordcount < count($englishByWord); $wordcount++) {
	                        array_push($items, new RenderItem(RenderTypes::eText, [new RenderText(RenderDisplayClass::eArabic, ArabicData::TransliterateFromBuckwalter($arabicText[$wordcount])), new RenderText(RenderDisplayClass::eTransliteration, $transliteration[$wordcount]), new RenderText(RenderDisplayClass::eLTR, $englishByWord[$wordcount])]));
	                    }
	                    array_push($renderer->Items, new RenderItem(RenderTypes::eText, [new RenderText(RenderDisplayClass::eNested, $items), new RenderText(RenderDisplayClass::eArabic, ArabicData::TransliterateFromBuckwalter($matches[$matchcount][1])), new RenderText(RenderDisplayClass::eTransliteration, Arabic::TransliterateToScheme(ArabicData::TransliterateFromBuckwalter($matches[$matchcount][1]), $schemetype, $scheme)), new RenderText(RenderDisplayClass::eLTR, Utility::LoadResourceString("IslamInfo_" . $translationID . "Trans"))]));
	                }
	                if (array_key_exists(3, $matches[$matchcount])) {
	                    $renderer->Items = array_merge($renderer->Items, DocBuilder::TextFromReferences($matches[$matchcount][3], $schemetype, $scheme, $translationindex)->Items);
	                }
	            }
			}
		}
		return $renderer;
	}
    public static function NormalTextFromReferences($ID, $strings, $schemetype, $scheme, $translationindex)
    {
        $renderer = new RenderArray($ID);
        if ($strings == null) { return $renderer; }
        if (preg_match_all("/(.*?)(?:(\\{)(.*?)(\\})|$)/u", $strings, $matches, PREG_SET_ORDER) !== 0) {
	        for ($count = 0; $count < count($matches); $count++) {
	            if (count($matches[$count]) != 0) {
	                if (strlen($matches[$count][1]) != 0) {
	                	array_push($renderer->Items, new RenderItem(RenderTypes::eText, [new RenderText(RenderDisplayClass::eLTR, $matches[$count][1])]));
	            	}
	                if (strlen($matches[$count][3]) != 0) {
	                    $renderer->Items = array_merge($renderer->Items, DocBuilder::TextFromReferences($matches[$count][3], $schemetype, $scheme, $translationindex)->Items);
	                }
	            }
	        }
    	}
        return $renderer;
    }
	public static function TextFromReferences($strings, $schemetype, $scheme, $translationindex)
	{
		$options = explode(";", $strings);
		$strings = array_shift($options);
		foreach ($options as $idx => $str) {
			$kv = explode("=", $str);
			$options[$kv[0]] = (count($kv) == 1) ? null : explode(",", $kv[1]);
			unset($options[$idx]);
		}
		if (array_key_exists("Translation", $options)) { $translationindex = TanzilReader::GetTranslationIndex($options["Translation"][0]); }
		if (array_key_exists("TranslitStyle", $options)) {
			if ($options["TranslitStyle"][0] == "Literal") {
				$schemetype = TranslitScheme::Literal;
			} elseif ($options["TranslitStyle"][0] == "RuleBased") {
				$schemetype = TranslitScheme::RuleBased;
			} else {
				$schemetype = TranslitScheme::None;
			}
		}
		if (array_key_exists("TranslitScheme", $options)) { $scheme = $options["TranslitScheme"][0]; }
		$renderer = new RenderArray("");
		if (TanzilReader::IsQuranTextReference($strings)) {
            $renderer->Items = array_merge($renderer->Items, TanzilReader::QuranTextFromReference($strings, $schemetype, $scheme, $translationindex)->Items);
        }
		for ($count = 0; $count < count(CachedData::IslamData()->abbreviations->children()); $count++) {
			for ($subcount = 0; $subcount < count(CachedData::IslamData()->abbreviations->children()[$count]->children()); $subcount++) {
				if (array_search($strings, explode("|", (string)CachedData::IslamData()->abbreviations->children()[$count]->children()[$subcount]->attributes()["text"])) !== false) {
					$items = array();
					if (array_key_exists("Char", $options)) $options["Char"] = array_map(function ($str) { return substr("00000000" . strtolower($str), -8); }, $options["Char"]);
					if (count($options) == 0) array_push($items, new RenderItem(RenderTypes::eText, [new RenderText(RenderDisplayClass::eTag, (string)CachedData::IslamData()->abbreviations->children()[$count]->children()[$subcount]->attributes()["id"] . "|" . (string)CachedData::IslamData()->abbreviations->children()[$count]->children()[$subcount]->attributes()["text"])]));
					if ((string)CachedData::IslamData()->abbreviations->children()[$count]->children()[$subcount]->attributes()["font"] != "") {
						foreach (explode("|", (string)CachedData::IslamData()->abbreviations->children()[$count]->children()[$subcount]->attributes()["font"]) as $part) {
							$font = "";
							if (strpos($part, ';') !== false) {
								$font = explode(';', $part)[0];
								$part = explode(';', $part)[1];
							}
							if (!array_key_exists("Font", $options) || array_search($font, $options["Font"]) !== false) {
								foreach (explode(",", $part) as $substr) {
									if (!array_key_exists("Char", $options) || array_search(bin2hex(mb_convert_encoding(implode(array_map(function($split) { return mb_convert_encoding('&#x' . $split . ';', 'UTF-8', 'HTML-ENTITIES'); }, explode("+", $substr))), 'UCS-4BE', 'UTF-8')), $options["Char"]) !== false) {
		                                $rendText = new RenderText(RenderDisplayClass::eArabic, implode(array_map(function($split) { return mb_convert_encoding('&#x' . $split . ';', 'UTF-8', 'HTML-ENTITIES'); }, explode("+", $substr))));
		                                $rendText->font = $font;
		                                array_push($items, new RenderItem(RenderTypes::eText, [$rendText]));
		                            }
								}
							}
						}
					}
					array_push($renderer->Items, new RenderItem(RenderTypes::eText, [new RenderText(RenderDisplayClass::eNested, $items)]));
					break;
				}
			}
		}
		return $renderer;
	}
};
function getcacheitem()
{
	//return false; //disable for debugging
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
		$font = dirname(__FILE__) . '/files/' . (array_key_exists("Font", $_GET) && array_search($_GET["Font"], Utility::$FontList) !== false ? Utility::$FontFile[array_search($_GET["Font"], Utility::$FontList)] : 'me_quran.ttf');
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
			$out = "";
			$renderer = DocBuilder::BuckwalterTextFromReferences(null, null, null, CachedData::IslamData()->lists->category->children()[$count]["text"], null, 0);
			for ($matchcount = 0; $matchcount < count($renderer->Items); $matchcount++) {
				$val = "";
				$id = "";
				for ($textcount = 0; $textcount < count($renderer->Items[$matchcount]->textitems); $textcount++) {
					if ($renderer->Items[$matchcount]->textitems[$textcount]->displayClass == RenderDisplayClass::eNested) {
						for ($nestcount = 0; $nestcount < count($renderer->Items[$matchcount]->textitems[$textcount]->Text); $nestcount++) {
							for ($nesttextcount = 0; $nesttextcount < count($renderer->Items[$matchcount]->textitems[$textcount]->Text[$nestcount]->textitems); $nesttextcount++) {
								if ($renderer->Items[$matchcount]->textitems[$textcount]->Text[$nestcount]->textitems[$nesttextcount]->font != "") {
									$val .= "{\"value\":" . json_encode(explode("|", $id)[1]) . ", \"font\": \"" . $renderer->Items[$matchcount]->textitems[$textcount]->Text[$nestcount]->textitems[$nesttextcount]->font . "\", \"char\": \"" . bin2hex(mb_convert_encoding($renderer->Items[$matchcount]->textitems[$textcount]->Text[$nestcount]->textitems[$nesttextcount]->Text, 'UCS-4BE', 'UTF-8')) . "\"},";
								} else if ($renderer->Items[$matchcount]->textitems[$textcount]->Text[$nestcount]->textitems[$nesttextcount]->displayClass == RenderDisplayClass::eTag) {
									$id = $renderer->Items[$matchcount]->textitems[$textcount]->Text[$nestcount]->textitems[$nesttextcount]->Text;
								}
							}
						}
					}
				}
				if ($val !== "") {
					$val = "{\"text\": \"" . Utility::LoadResourceString("IslamInfo_" . explode("|", $id)[0]) . "\", \"values\":[" . substr($val, 0, -1) . "]}";
				}
				if ($val !== "") $out .= (($out !== "") ? "," : "") . $val;
			}
			if ($out !== "") $json .= (($json !== "") ? "," : "") . "{\"text\": \"" . Utility::LoadResourceString("IslamInfo_" . (string)CachedData::IslamData()->lists->category->children()[$count]["id"]) . "\", \"menu\":[" . $out . "]}";
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
   	$plugin_array['is_button'] = $url.'/islamsource/isbutton.js';
  	return $plugin_array;
}

/* REPLACE SHORTCODES */

function is_shortcode() {
    global $wp_query;	
    $posts = $wp_query->posts;
    foreach ($posts as $post){
		$post->post_content = RenderArray::DoRender(DocBuilder::NormalTextFromReferences(null, $post->post_content, null, null, 0)->Items);
    }
}
add_action( 'wp', 'is_shortcode' );

}
?>