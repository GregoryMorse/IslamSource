<?php
/*
Plugin Name: IslamSource
Description: Islam Source Phrase Plugin
Version: 1.0.0
Author: IslamSource
Author URI: http://islamsource.info

*/

/* OPTIONS PAGE */

// Add admin actions
add_action('admin_init', 'islamic_source_init');

// Register settings
function islamic_source_init(){
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

class CachedData
{
	public static $_XMLDocMain = null;
	public static function XMLDocMain() { if ($_XMLDocMain === null) $_XMLDocMain = simplexml_load_file(dirname(__FILE__) . "\\metadata\\" . TanzilReader::$QuranTextNames[0] . ".xml"); return $_XMLDocMain; }
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
	public static function GetTextChapter($xmldoc, $chapter)
	{
		return TanzilReader::GetChildNodeByIndex("sura", "index", $chapter, $xmldoc->children());
	}
	public static function GetTextVerse($chapternode, $verse)
	{
		return TanzilReader::GetChildNodeByIndex("aya", "index", $verse, $chapternode->children());
	}
	public static function GetChapterCount()
	{
        return TanzilReader::GetChildNodeCount("sura", CachedData.XMLDocInfo()->children()["suras"]);
    }
    public static $QuranTextNames = array("quran-hafs", "quran-warsh", "quran-alduri");
};

/* REPLACE SHORTCODES */

function is_shortcode() {
    global $wp_query;	
    $posts = $wp_query->posts;
    foreach ($posts as $post){
		$post->post_content = preg_replace_callback('/(.*?)(?:(\{)(.*?)(\})|$)/s', function ($matches) {
			if (TanzilReader::IsQuranTextReference($matches[3])) {
				return $matches[1] . TanzilReader::QuranTextFromReference($matches[3]);
			} else {
				return $matches[0];
			}
		}, $post->post_content);
    }
}
add_action( 'wp', 'is_shortcode' );

/* ISLAMIC GRAPHIC SHORTCODE FUNCTION */

function insert_islamic_graphic( $atts, $content=null, $code="" ) {
        $code = preg_replace("/_w/", "", $code); // to be compatible with pre-v1.1 shortcodes   
        
        if ($display_type == 'images' || $display_type == 'images_trans'){

            // extract attributes
            extract( shortcode_atts( array(
                    'h' => $default_height,
                    'c' => $default_colour,
            ), $atts ) );
            
            // plugin URL
            $plugin_url = plugin_dir_url( "IslamSourceWP.php" );
            
            // Construct alt_text
            $alt_text_str =  $romanized . ' (' . $translation . ')';

            // Build HTML
            $html  = '<img title="' . $alt_text_str . '"';
            $html .= ' alt="' . $alt_text_str . '"';
            $html .= ' class="islamic_graphic"';
            $html .= ' src="'. $plugin_url . 'islam-source/img/' . "{$c}" . '/';

            if ("{$h}" <= 20) { $html .= "20"; }
            else { $html .= "40"; }

            $html .= '/' . $code . '.png" height="'. "{$h}" . 'px">';
            
            if ($display_type == 'images_trans') {
                $html .= '<span class="islamic_graphic"> (' . $translation . ')</span>';
            }
            
        }       
        elseif ($display_type == 'text_rom_trans') {          
            $html = '<span class="islamic_graphic"> '. $romanized . ' (' . $translation . ')</span>';           
        }
        elseif ($display_type == 'text_rom') {          
            $html = '<span class="islamic_graphic"> ' . $romanized . '</span>';         
        }
        elseif ($display_type == 'text_trans') {          
            $html = '<span class="islamic_graphic">(' . $translation . ')</span>';         
        }
        
	return $html;
	}

	//Handle an individual file import.
	function handle_import_file($file, $post_id = 0, $import_date = 'file') {
		set_time_limit(120);
		
		// Initially, Base it on the -current- time.
		$time = current_time('mysql', 1);
		// Next, If it's post to base the upload off:
		if ( 'post' == $import_date && $post_id > 0 ) {
			$post = get_post($post_id);
			if ( $post && substr( $post->post_date_gmt, 0, 4 ) > 0 )
				$time = $post->post_date_gmt;
		} elseif ( 'file' == $import_date ) {
			$time = gmdate( 'Y-m-d H:i:s', @filemtime($file) );
		}

		// A writable uploads dir will pass this test. Again, there's no point overriding this one.
		if ( ! ( ( $uploads = wp_upload_dir($time) ) && false === $uploads['error'] ) )
			return new WP_Error( 'upload_error', $uploads['error']);

		$wp_filetype = wp_check_filetype( $file, null );

		extract( $wp_filetype );
		
		if ( ( !$type || !$ext ) && !current_user_can( 'unfiltered_upload' ) )
			return new WP_Error('wrong_file_type', __( 'Sorry, this file type is not permitted for security reasons.' ) ); //A WP-core string..

		//Is the file allready in the uploads folder?
		if ( preg_match('|^' . preg_quote(str_replace('\\', '/', $uploads['basedir'])) . '(.*)$|i', $file, $mat) ) {

			$filename = basename($file);
			$new_file = $file;

			$url = $uploads['baseurl'] . $mat[1];

			$attachment = get_posts(array( 'post_type' => 'attachment', 'meta_key' => '_wp_attached_file', 'meta_value' => ltrim($mat[1], '/') ));
			if ( !empty($attachment) )
				return new WP_Error('file_exists', __( 'Sorry, That file already exists in the WordPress media library.', 'add-from-server' ) );

			//Ok, Its in the uploads folder, But NOT in WordPress's media library.
			if ( 'file' == $import_date ) {
				$time = @filemtime($file);
				if ( preg_match("|(\d+)/(\d+)|", $mat[1], $datemat) ) { //So lets set the date of the import to the date folder its in, IF its in a date folder.
					$hour = $min = $sec = 0;
					$day = 1;
					$year = $datemat[1];
					$month = $datemat[2];
	
					// If the files datetime is set, and it's in the same region of upload directory, set the minute details to that too, else, override it.
					if ( $time && date('Y-m', $time) == "$year-$month" )
						list($hour, $min, $sec, $day) = explode(';', date('H;i;s;j', $time) );
	
					$time = mktime($hour, $min, $sec, $month, $day, $year);
				}
				$time = gmdate( 'Y-m-d H:i:s', $time);
				
				// A new time has been found! Get the new uploads folder:
				// A writable uploads dir will pass this test. Again, there's no point overriding this one.
				if ( ! ( ( $uploads = wp_upload_dir($time) ) && false === $uploads['error'] ) )
					return new WP_Error( 'upload_error', $uploads['error']);
				$url = $uploads['baseurl'] . $mat[1];
			}
		} else {
			$filename = wp_unique_filename( $uploads['path'], basename($file));

			// copy the file to the uploads dir
			$new_file = $uploads['path'] . '/' . $filename;
			if ( false === @copy( $file, $new_file ) )
				return new WP_Error('upload_error', sprintf( __('The selected file could not be copied to %s.', 'add-from-server'), $uploads['path']) );

			// Set correct file permissions
			$stat = stat( dirname( $new_file ));
			$perms = $stat['mode'] & 0000666;
			@ chmod( $new_file, $perms );
			// Compute the URL
			$url = $uploads['url'] . '/' . $filename;
			
			if ( 'file' == $import_date )
				$time = gmdate( 'Y-m-d H:i:s', @filemtime($file));
		}

		//Apply upload filters
		$return = apply_filters( 'wp_handle_upload', array( 'file' => $new_file, 'url' => $url, 'type' => $type ) );
		$new_file = $return['file'];
		$url = $return['url'];
		$type = $return['type'];

		$title = preg_replace('!\.[^.]+$!', '', basename($file));
		$content = '';

		// use image exif/iptc data for title and caption defaults if possible
		if ( $image_meta = @wp_read_image_metadata($new_file) ) {
			if ( '' != trim($image_meta['title']) )
				$title = trim($image_meta['title']);
			if ( '' != trim($image_meta['caption']) )
				$content = trim($image_meta['caption']);
		}

		if ( $time ) {
			$post_date_gmt = $time;
			$post_date = $time;
		} else {
			$post_date = current_time('mysql');
			$post_date_gmt = current_time('mysql', 1);
		}

		// Construct the attachment array
		$attachment = array(
			'post_mime_type' => $type,
			'guid' => $url,
			'post_parent' => $post_id,
			'post_title' => $title,
			'post_name' => $title,
			'post_content' => $content,
			'post_date' => $post_date,
			'post_date_gmt' => $post_date_gmt
		);

		$attachment = apply_filters('afs-import_details', $attachment, $file, $post_id, $import_date);

		//Win32 fix:
		$new_file = str_replace( strtolower(str_replace('\\', '/', $uploads['basedir'])), $uploads['basedir'], $new_file);

		// Save the data
		$id = wp_insert_attachment($attachment, $new_file, $post_id);
		if ( !is_wp_error($id) ) {
			$data = wp_generate_attachment_metadata( $id, $new_file );
			wp_update_attachment_metadata( $id, $data );
		}
		//update_post_meta( $id, '_wp_attached_file', $uploads['subdir'] . '/' . $filename );

		return $id;
	}

//simplexml_load_string
//metadata/islaminfo.xml
//metadata/ArabicData.xml
//IslamResources/My Project/Resources.resx, IslamResources/Resources.en.resx, etc...
?>