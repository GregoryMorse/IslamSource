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

/* ADD SHORTCODES */

function is_shortcode() {
    global $wp_query;	
    $posts = $wp_query->posts;
    foreach ($posts as $post){
		$post->post_content = preg_replace_callback( '/(?:,?(\d+)(?:\:(\d+))?(?:\:(\d+))?(?:-(\d+)(?:\:(\d+))?(?:\:(\d+))?)?)/s', function () {
			return "<span>Quran Reference</span>";
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