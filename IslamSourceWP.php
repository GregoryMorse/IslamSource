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
add_action('admin_menu', 'islamic_graphics_menu');
add_action('admin_init', 'islamic_graphics_init');

// Add sub-menu to Settings Top Menu
function islamic_graphics_menu() {
	add_options_page('Islamic Graphics - Options', 'Islamic Graphics', 'manage_options', 'islamic-graphics', 'islamic_graphics_options');
}

// Register settings
function islamic_graphics_init(){
    add_settings_section('islamic_graphics_main', 'Main Settings', 'islamic_graphics_section_text', 'islamic-graphics');
    add_settings_field('islamic_graphics_default_height', 'Default height (pixels)', 'add_field_default_height', 'islamic-graphics', 'islamic_graphics_main');
    add_settings_field('islamic_graphics_default_colour', 'Default colour', 'add_field_default_colour', 'islamic-graphics', 'islamic_graphics_main');
    add_settings_field('islamic_graphics_display_type', 'Display type', 'add_field_display_type', 'islamic-graphics', 'islamic_graphics_main');
    
    register_setting( 'islamic-graphics-option-group', 'islamic_graphics_default_height', 'validate_default_height' );
    register_setting( 'islamic-graphics-option-group', 'islamic_graphics_display_type' );
    register_setting( 'islamic-graphics-option-group', 'islamic_graphics_default_colour' );
 }
 
function islamic_graphics_section_text(){
    echo '<p>Use the following options to alter the display of Islamic Graphics in your posts and pages.</p>';
}

function add_field_default_height(){   
    echo '<input type="text" name="islamic_graphics_default_height" value="';
    echo get_option('islamic_graphics_default_height', 20);
    echo '" />';           
}

function add_field_default_colour(){   
    echo '<select name="islamic_graphics_default_colour" id="islamic_graphics_default_colour">';
    
    $display_type_options = array(
            "black" => "Black",
            "white" => "White");

    $stored_type = get_option('islamic_graphics_default_colour', 'black');

    foreach ($display_type_options as $key => $row) {
        echo '<option value="' . $key . '"';
        if ($stored_type == $key) { echo 'selected="selected"'; }
        echo '>'. $row .'</option>';
    }
    
    echo '</select>';        
}

function add_field_display_type(){   
    echo '<select name="islamic_graphics_display_type" id="islamic_graphics_display_type">';
    
        $display_type_options = array(
                "images" => "Images only",
                "images_trans" => "Images (with translation)",
                "text_rom_trans" => "Romanized text (with translation)",
                "text_rom" => "Romanized text only",
                "text_trans" => "Translation only");

        $stored_type = get_option('islamic_graphics_display_type', 'images');

        foreach ($display_type_options as $key => $row) {
            echo '<option value="' . $key . '"';
            if ($stored_type == $key) { echo 'selected="selected"'; }
            echo '>'. $row .'</option>';
        } 
    
    echo '</select>';        
}

// Validate height input
function validate_default_height($input) {
    $newinput = trim($input);
    
    if (!is_numeric($newinput)){
        $newinput = 20; // if not numeric, then use 20 as default value
    }
    return $newinput;
}

// HTML for options page
function islamic_graphics_options() {
	if (!current_user_can('manage_options'))  {
		wp_die( __('You do not have sufficient permissions to access this page.') );
	}
?>
    <div class="wrap">
        <h2>Islamic Graphics - Options Page</h2>
        <form method="post" action="options.php">
        <?php settings_fields( 'islamic-graphics-option-group' ); ?>
        <?php do_settings_sections('islamic-graphics'); ?>
        <!-- Submit Button -->
        <p class="submit"> <input type="submit" class="button-primary" value="<?php _e('Save Changes') ?>" /></p>
        </form>
    </div>
        
<?php }



/* GLOBAL ARRAY OF SHORTCODES + ALT TEXT VALUES */

$alt_text = array();

function add_to_alt_text($code, $romanized, $translation){
    $GLOBALS['alt_text'][$code] = array("romanized" => $romanized, "translation" => $translation); 
}

/*add_to_alt_text("alayhis", "'alayhi'l-salam", "peace be upon him");
add_to_alt_text("rahimaha", "ra?imaha Allah", "may Allah have mercy upon her");
add_to_alt_text("rahimahu", "ra?imahullah","may Allah have mercy upon him");
add_to_alt_text("rahimahum", "ra?imahum Allah","may Allah have mercy upon them");
add_to_alt_text("ranha", "ra?yAllahu 'anha","may Allah be pleased with her");
add_to_alt_text("ranhu", "ra?yAllahu 'anhu","may Allah be pleased with him");
add_to_alt_text("ranhum", "ra?yAllahu 'anhum","may Allah be pleased with them");
add_to_alt_text("saw", "?allallahu 'alayhi wa sallam","peace and blessings of Allah be upon him");
add_to_alt_text("swt", "sub?anahu wa ta'ala","glorified and exalted be He");
*/

/* ADD SHORTCODES */
        
foreach ($alt_text as $key => $row) {
    add_shortcode( $key, 'insert_islamic_graphic' );
    add_shortcode( $key.'_w', 'insert_islamic_graphic' ); // to be compatible with pre-v1.1 shortcodes
}   


/* ISLAMIC GRAPHIC SHORTCODE FUNCTION */

function insert_islamic_graphic( $atts, $content=null, $code="" ) {
        $code = preg_replace("/_w/", "", $code); // to be compatible with pre-v1.1 shortcodes   
    
        // fetch display type from wp options (default to images)
        $display_type = get_option('islamic_graphics_display_type', 'images');
        
        // Fetch alt_text
        $romanized = $GLOBALS['alt_text'][$code]['romanized'];
        $translation = $GLOBALS['alt_text'][$code]['translation'];
        
        if ($display_type == 'images' || $display_type == 'images_trans'){
            // fetch default height from wp options (default of 20 if not set)
            $default_height = get_option('islamic_graphics_default_height', 20);
            
            // fetch default colour from wp_options (default to black if not set)
            $default_colour = get_option('islamic_graphics_default_colour', 'black');

            // extract attributes
            extract( shortcode_atts( array(
                    'h' => $default_height,
                    'c' => $default_colour,
            ), $atts ) );
            
            // plugin URL
            $plugin_url = plugin_dir_url( "islamic_graphics.php" );
            
            // Construct alt_text
            $alt_text_str =  $romanized . ' (' . $translation . ')';

            // Build HTML
            $html  = '<img title="' . $alt_text_str . '"';
            $html .= ' alt="' . $alt_text_str . '"';
            $html .= ' class="islamic_graphic"';
            $html .= ' src="'. $plugin_url . 'islamic-graphics/img/' . "{$c}" . '/';

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

//GD library and FreeType library
//header('Content-Type: image/png');
//$font = './files/.ttf';
//$bbox = imagettfbbox(0, 0, $font, $text);
//$im = imagecreatetruecolor($bbox[4] - $bbox[0], $bbox[5] - $bbox[1]);
//$white = imagecolorallocate($im, 255, 255, 255);
//imagefilledrectangle($im, $bbox[0], $bbox[1], $bbox[4], $bbox[5], $white);
//imagettftext($im, $bbox[0], $bbox[1], $bbox[4], $bbox[5], $black, $font, $text);
//imagepng($im);
//imagedestroy($im);
?>