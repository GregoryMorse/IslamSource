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

class CachedData
{
	public static $_IslamData = null;
	public static function IslamData() { if (CachedData::$_IslamData === null) CachedData::$_IslamData = simplexml_load_file(dirname(__FILE__) . "/metadata/islaminfo.xml"); return CachedData::$_IslamData; }
	public static $_XMLDocMain = null;
	public static function XMLDocMain() { if ($_XMLDocMain === null) $_XMLDocMain = simplexml_load_file(dirname(__FILE__) . "/metadata/" . TanzilReader::$QuranTextNames[0] . ".xml"); return $_XMLDocMain; }
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
	public static function DoGetRenderedQuranText($qurantext, $basechapter, $baseverse, $translation, $schemetype, $scheme, $translationindex)
	{
		$text = "";
		if ($qurantext !== null) {
			for ($chapter = 0; $chapter < count($qurantext); $chapter++) {
				for ($verse = 0; $verse < count($qurantext[$chapter]); $verse++) {
					$text .= $qurantext[$chapter][$verse];
				}
			}
		}
		return $text;
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

function TextFromReferences($str)
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
			$matches = array();
			$out = "";
			if (preg_match_all('/(.*?)(?:(\\\{)(.*?)(\\\})|$)/s', CachedData::IslamData()->lists->category->children()[$count]["text"], $matches, PREG_SET_ORDER) !== 0) {
				$val = "";
				for ($matchcount = 0; $matchcount < count($matches); $matchcount++) {
					if (count($matches[$matchcount]) >= 4) {
						$check = TextFromReferences($matches[$matchcount][3]);
						if ($check !== "") $val .= (($val !== "") ? "," : "") . $check;
					}
				}
				if ($val !== "") $out .= (($out !== "") ? "," : "") . "{\"text\": \"" . Utility::LoadResourceString("IslamInfo_" . (string)CachedData::IslamData()->lists->category->children()[$count]["id"]) . "\", \"menu\":[" . $val . "]}";
			}
			$json .= (($json !== "") ? "," : "") . $out;
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