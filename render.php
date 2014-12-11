<?php
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
if (!array_key_exists("Cat", $_GET)) {
//GD library and FreeType library
$size = $_GET["Size"];
if ($size === null) $size = 32;
$chr = mb_convert_encoding("&#" . hexdec($_GET["Char"]) . ";", 'UTF-8', 'HTML-ENTITIES');
$fonts = array("AGAIslamicPhrases", "AGAArabesque", "Shia", "IslamicLogo", "KFGQPCArabicSymbols01", "Quranic", "Tulth", "Farsi", "Asmaul-Husna", "Asmaul-Husna_2");
$fontfiles = array("AGA_Islamic_Phrases.TTF", "aga-arabesque.ttf", "SHIA.TTF", "islamic.ttf", "Symbols1_Ver02.otf", "Quranic.ttf", "Tulth.ttf", "Farsi.ttf", "Asmaul-Husna_1.ttf", "Asmaul-Husna_2.ttf");
$font = dirname(__FILE__) . '\\files\\' . (array_key_exists("Font", $_GET) && array_search($_GET["Font"], $fonts) !== false ? $fontfiles[array_search($_GET["Font"], $fonts)] : 'me_quran.ttf');
$bbox = imageftbbox($size, 0, $font, $chr);
$im = imagecreatetruecolor(ceil($bbox[4] - $bbox[0]), ceil($bbox[1] - $bbox[5]));
$white = imagecolorallocate($im, 255, 255, 255);
$black = imagecolorallocate($im, 0, 0, 0);
imagefilledrectangle($im, 0, 0, ceil($bbox[4] - $bbox[0]), ceil($bbox[1] - $bbox[5]), $white);
imagefttext($im, $size, 0, $bbox[0], -$bbox[5], $black, $font, $chr);
header('Content-Type: image/png');
imagepng($im);
imagedestroy($im);
} else {


class CachedData
{
	public static $_IslamData = null;
	public static function IslamData() { if (CachedData::$_IslamData === null) CachedData::$_IslamData = simplexml_load_file(dirname(__FILE__) . "\\metadata\\islaminfo.xml"); return CachedData::$_IslamData; }
};

class Utility
{
	public static $_resxml = null;
	public static function LoadResourceString($resourcekey)
	{
		if (Utility::$_resxml === null) Utility::$_resxml = simplexml_load_file(dirname(__FILE__) . "\\IslamResources\\My Project\\Resources.resx");
		$arr = Utility::$_resxml->xpath("/root/data[@name='" . $resourcekey . "']");
		if (count($arr) !== 0) return (string)$arr[0]->value;
		return null;
	}
}

function TextFromReferences($str)
{
	$val = "null";
	for ($count = 0; $count < count(CachedData::IslamData()->abbreviations->children()); $count++) {
		for ($subcount = 0; $subcount < count(CachedData::IslamData()->abbreviations->children()[$count]->children()); $subcount++) {
			if (array_search($str, explode("|", (string)CachedData::IslamData()->abbreviations->children()[$count]->children()[$subcount]->attributes()["text"])) !== false) {
				$val = "{\"text\": \"" . Utility::LoadResourceString("IslamInfo_" . (string)CachedData::IslamData()->abbreviations->children()[$count]->children()[$subcount]->attributes()["id"]) . "\", \"values\":[";
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
				$val .= "null]}";
				break;
			}
		}
	}
	return $val;
}
echo "[";
for ($count = 0; $count < count(CachedData::IslamData()->lists->category->children()); $count++) {
	$matches = array();
	if (preg_match_all('/(.*?)(?:(\\\{)(.*?)(\\\})|$)/s', CachedData::IslamData()->lists->category->children()[$count]["text"], $matches, PREG_SET_ORDER) !== 0) {
		if ($count !== 0) echo ",";
		echo "{\"text\": \"" . Utility::LoadResourceString("IslamInfo_" . (string)CachedData::IslamData()->lists->category->children()[$count]["id"]) . "\", \"menu\":[";
		for ($matchcount = 0; $matchcount < count($matches); $matchcount++) {
			if (count($matches[$matchcount]) >= 4) {
				if ($matchcount !== 0) echo ",";
				echo TextFromReferences($matches[$matchcount][3]);
			}
		}
		echo "]}";
	}
}
echo "]";
}
?>